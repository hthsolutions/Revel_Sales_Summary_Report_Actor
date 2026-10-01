import { readFile } from 'node:fs/promises';

import { Actor, log } from 'apify';
import { PlaywrightCrawler } from 'crawlee';
import * as XLSX from 'xlsx';
import { createClient } from '@supabase/supabase-js';

await Actor.init();

const supabaseUrl =
    process.env.SUPABASE_URL
    || 'https://ongqhvokcwceqgnetonq.supabase.co';

const supabaseServiceRoleKey =
    process.env.SUPABASE_SERVICE_ROLE_KEY;

if (!supabaseServiceRoleKey) {
    throw new Error(
        'SUPABASE_SERVICE_ROLE_KEY is not configured.',
    );
}

const supabase = createClient(
    supabaseUrl,
    supabaseServiceRoleKey,
    {
        auth: {
            persistSession: false,
            autoRefreshToken: false,
        },
    },
);

/**
 * Capture a screenshot and save it in the run's
 * default key-value store.
 */
async function saveScreenshot(page, key) {
    const screenshot = await page.screenshot({
        fullPage: true,
    });

    await Actor.setValue(key, screenshot, {
        contentType: 'image/png',
    });

    log.info(`Saved screenshot: ${key}`);
}

/**
 * Wait until Revel has finished rebuilding the report.
 *
 * Hidden loaders can remain in Revel's DOM permanently, so check whether any
 * matching loader is actually visible instead of waiting for the elements to
 * be detached.
 */
async function waitForReportIdle(page) {
    await page.waitForFunction(
        () => {
            const reportArea =
                document.querySelector('.report-content')
                ?? document.querySelector('.report-container')
                ?? document.querySelector('.reports-content')
                ?? document.body;

            const loadingElements = reportArea.querySelectorAll([
                '.loading',
                '.loader',
                '.spinner',
                '.loading-mask',
                '.blockUI',
                '.fa-spinner',
                '.icon-spinner',
                '[class*="loading-indicator"]',
            ].join(','));

            return [...loadingElements].every((element) => {
                const style = window.getComputedStyle(element);
                const bounds = element.getBoundingClientRect();

                return (
                    style.display === 'none'
                    || style.visibility === 'hidden'
                    || style.opacity === '0'
                    || bounds.width === 0
                    || bounds.height === 0
                );
            });
        },
        undefined,
        {
            timeout: 90_000,
            polling: 500,
        },
    );

    /*
     * Give computed totals and charts a brief opportunity to settle
     * after the loading indicator disappears.
     */
    await page.waitForTimeout(1_500);
}

function validateDate(value, fieldName) {
    const datePattern =
        /^(0?[1-9]|1[0-2])\/(0?[1-9]|[12]\d|3[01])\/\d{4}$/;

    if (!datePattern.test(value)) {
        throw new Error(
            `${fieldName} must use MM/DD/YYYY format. `
            + `Received: ${value}`,
        );
    }
}

function validateTime(value, fieldName) {
    const timePattern = /^(0?[1-9]|1[0-2]):[0-5]\d$/;

    if (!timePattern.test(value)) {
        throw new Error(
            `${fieldName} must use HH:MM 12-hour format. `
            + `Received: ${value}`,
        );
    }
}

const CENTRAL_TIME_ZONE = 'America/Chicago';

/**
 * Calendar date for a moment in US Central Time.
 */
function getCentralDateParts(date) {
    const parts = new Intl.DateTimeFormat('en-US', {
        timeZone: CENTRAL_TIME_ZONE,
        year: 'numeric',
        month: '2-digit',
        day: '2-digit',
    }).formatToParts(date);

    const lookup = Object.fromEntries(
        parts
            .filter((part) => part.type !== 'literal')
            .map((part) => [part.type, Number(part.value)]),
    );

    return {
        year: lookup.year,
        month: lookup.month,
        day: lookup.day,
    };
}

function formatDateParts({ year, month, day }) {
    const monthText = String(month).padStart(2, '0');
    const dayText = String(day).padStart(2, '0');

    return `${monthText}/${dayText}/${year}`;
}

/**
 * Report start date: yesterday in US Central Time, MM/DD/YYYY.
 */
function calculateStartDate(today = new Date()) {
    const { year, month, day } = getCentralDateParts(today);
    const calendarDate = new Date(Date.UTC(year, month - 1, day));

    calendarDate.setUTCDate(calendarDate.getUTCDate() - 1);

    return formatDateParts({
        year: calendarDate.getUTCFullYear(),
        month: calendarDate.getUTCMonth() + 1,
        day: calendarDate.getUTCDate(),
    });
}

/**
 * Report end date: today in US Central Time, MM/DD/YYYY.
 */
function calculateEndDate(today = new Date()) {
    return formatDateParts(getCentralDateParts(today));
}

const FILTER_COLUMNS_SET_TO_ALL = [
    'Dining Options',
    'Pos Stations',
    'Employees',
];

const FILTER_COLUMNS_CLEARED = [
    'Online orders',
    'Inclusions',
];

/**
 * Force the Sales Summary filter panel into a known state.
 *
 * Revel remembers the last filter selection for the account and ships with
 * Unpaid and Irregular pre-checked under Inclusions, so an unattended run can
 * otherwise export a differently scoped report than expected.
 */
async function applyReportFilters(page) {
    const filtersArea = page
        .locator('.sales-summary .filters-area')
        .first();

    const filtersToggle = filtersArea.locator('.filters-button').first();

    /*
     * The filter form itself can report as hidden even while the panel is
     * open, so use a column heading as the open/closed signal.
     */
    const panelHeading = filtersArea
        .getByText(FILTER_COLUMNS_SET_TO_ALL[0], { exact: true })
        .first();

    await filtersArea.waitFor({
        state: 'visible',
        timeout: 20_000,
    });

    /*
     * Switching establishments reloads the report, and the Filters button
     * ignores clicks until Revel has rebound its handlers.
     */
    await waitForReportIdle(page);

    const maxOpenAttempts = 3;

    for (
        let attempt = 1;
        !(await panelHeading.isVisible());
        attempt += 1
    ) {
        if (attempt > maxOpenAttempts) {
            throw new Error(
                `The Revel filters panel did not open after `
                + `${maxOpenAttempts} attempts.`,
            );
        }

        log.info(
            `Opening the report filters panel (attempt ${attempt}).`,
        );

        await filtersToggle.click();

        await panelHeading
            .waitFor({
                state: 'visible',
                timeout: 10_000,
            })
            .catch(() => {});
    }

    await saveScreenshot(page, 'REVEL_FILTERS_BEFORE');

    const filterState = await filtersArea.evaluate(
        (root, { allColumns, clearedColumns }) => {
            if (!root.querySelector('form#filter_form')) {
                throw new Error(
                    'The Revel filter form (#filter_form) was not found.',
                );
            }

            const normalize = (value) => String(value ?? '')
                .replace(/\s+/g, ' ')
                .trim();

            /*
             * Column headings are matched on their own text so that an outer
             * wrapper holding several columns can never be mistaken for one.
             */
            const ownText = (element) => normalize(
                [...element.childNodes]
                    .filter((node) => node.nodeType === Node.TEXT_NODE)
                    .map((node) => node.textContent)
                    .join(' '),
            );

            const titles = [...allColumns, ...clearedColumns];
            const headings = new Map();

            for (const element of root.querySelectorAll('*')) {
                const text = ownText(element).toLowerCase();

                const title = titles.find(
                    (candidate) => candidate.toLowerCase() === text,
                );

                if (title && !headings.has(title)) {
                    headings.set(title, element);
                }
            }

            const missingTitles = titles.filter(
                (title) => !headings.has(title),
            );

            if (missingTitles.length > 0) {
                throw new Error(
                    `Filter columns missing from the Revel filter panel: `
                    + `${missingTitles.join(', ')}`,
                );
            }

            const columnElements = new Map();

            for (const [title, heading] of headings) {
                const otherHeadings = [...headings.entries()]
                    .filter(([otherTitle]) => otherTitle !== title)
                    .map(([, element]) => element);

                let node = heading;

                while (
                    node
                    && node !== root
                    && !node.querySelector('input')
                ) {
                    node = node.parentElement;
                }

                if (
                    !node
                    || node === root
                    || otherHeadings.some((other) => node.contains(other))
                ) {
                    throw new Error(
                        `Unable to isolate the "${title}" filter column. `
                        + `The Revel filter markup has changed.`,
                    );
                }

                columnElements.set(title, node);
            }

            const labelFor = (input) => {
                if (input.id) {
                    const explicit = root.querySelector(
                        `label[for="${CSS.escape(input.id)}"]`,
                    );

                    if (explicit) return normalize(explicit.textContent);
                }

                const wrapper = input.closest('label');

                if (wrapper) return normalize(wrapper.textContent);

                return normalize(input.parentElement?.textContent)
                    || normalize(input.value)
                    || normalize(input.name)
                    || '(unlabelled)';
            };

            /*
             * Revel binds jQuery change handlers to these inputs, so route the
             * update through jQuery when it is available.
             */
            const setChecked = (input, shouldBeChecked) => {
                if (input.checked === shouldBeChecked) return false;

                if (window.jQuery) {
                    window.jQuery(input)
                        .prop('checked', shouldBeChecked)
                        .trigger('change');
                } else {
                    input.checked = shouldBeChecked;

                    input.dispatchEvent(
                        new Event('change', { bubbles: true }),
                    );
                }

                return true;
            };

            const changes = [];

            for (const title of allColumns) {
                const column = columnElements.get(title);

                const radios = [
                    ...column.querySelectorAll('input[type="radio"]'),
                ];

                const allRadio = radios.find(
                    (radio) => /^all\b/i.test(labelFor(radio)),
                );

                if (!allRadio) {
                    throw new Error(
                        `No "All" option found in the "${title}" `
                        + `filter column.`,
                    );
                }

                if (setChecked(allRadio, true)) {
                    changes.push(
                        `${title}: selected "${labelFor(allRadio)}"`,
                    );
                }

                for (const radio of radios) {
                    if (radio !== allRadio && radio.checked) {
                        setChecked(radio, false);

                        changes.push(
                            `${title}: cleared "${labelFor(radio)}"`,
                        );
                    }
                }
            }

            for (const title of clearedColumns) {
                const column = columnElements.get(title);

                const checkboxes = column.querySelectorAll(
                    'input[type="checkbox"]',
                );

                for (const checkbox of checkboxes) {
                    if (setChecked(checkbox, false)) {
                        changes.push(
                            `${title}: cleared "${labelFor(checkbox)}"`,
                        );
                    }
                }
            }

            /*
             * Re-read every input after the handlers have run so the returned
             * state reflects what Revel will actually submit.
             */
            const columns = {};

            for (const [title, column] of columnElements) {
                columns[title] = [
                    ...column.querySelectorAll(
                        'input[type="radio"], input[type="checkbox"]',
                    ),
                ].map((input) => ({
                    label: labelFor(input),
                    type: input.type,
                    checked: input.checked,
                }));
            }

            for (const title of allColumns) {
                const checked = columns[title].filter(
                    (input) => input.checked,
                );

                if (
                    checked.length !== 1
                    || !/^all\b/i.test(checked[0].label)
                ) {
                    throw new Error(
                        `The "${title}" filter column is not set to "All". `
                        + `Currently checked: `
                        + `${checked
                            .map((input) => input.label)
                            .join(', ') || 'nothing'}`,
                    );
                }
            }

            for (const title of clearedColumns) {
                const checked = columns[title].filter(
                    (input) => input.checked,
                );

                if (checked.length > 0) {
                    throw new Error(
                        `The "${title}" filter column still has `
                        + `selections: `
                        + `${checked
                            .map((input) => input.label)
                            .join(', ')}`,
                    );
                }
            }

            return { columns, changes };
        },
        {
            allColumns: FILTER_COLUMNS_SET_TO_ALL,
            clearedColumns: FILTER_COLUMNS_CLEARED,
        },
    );

    await Actor.setValue('REVEL_FILTER_STATE', filterState);

    for (const [title, inputs] of Object.entries(filterState.columns)) {
        const checked = inputs
            .filter((input) => input.checked)
            .map((input) => input.label);

        log.info(
            `Filter "${title}": `
            + `${checked.join(', ') || 'nothing selected'}`,
        );
    }

    if (filterState.changes.length === 0) {
        log.info('Report filters already matched the required state.');
    } else {
        log.info(
            `Report filters adjusted: `
            + `${filterState.changes.join('; ')}`,
        );
    }

    await saveScreenshot(page, 'REVEL_FILTERS_SET');

    let applyButton = filtersArea
        .locator('button, a, .button')
        .filter({ hasText: /^\s*Apply\s*$/ })
        .first();

    if (await applyButton.count() === 0) {
        applyButton = filtersArea
            .locator('input[type="submit"][value="Apply"]')
            .first();
    }

    await applyButton.waitFor({
        state: 'visible',
        timeout: 20_000,
    });

    log.info('Applying the report filters.');

    await applyButton.click();

    await panelHeading
        .waitFor({
            state: 'hidden',
            timeout: 30_000,
        })
        .catch(async () => {
            log.info(
                'The filters panel stayed open after Apply. Closing it.',
            );

            await filtersToggle.click();

            await panelHeading.waitFor({
                state: 'hidden',
                timeout: 15_000,
            });
        });

    await waitForReportIdle(page);

    const filterSummary = (
        await filtersArea
            .locator('.chosen.options')
            .first()
            .innerText()
            .catch(() => '')
    ).replace(/\s+/g, ' ').trim();

    log.info(
        `Filter summary shown by Revel: `
        + `${filterSummary || 'none'}`,
    );

    if (filterSummary !== '') {
        log.warning(
            `Revel still reports active filters after Apply: `
            + `${filterSummary}`,
        );
    }

    await saveScreenshot(page, 'REVEL_FILTERS_APPLIED');
}

/**
 * Convert each worksheet into a nested dictionary by mapping
 * row 1 headers to row 2 values. Row 3 contains report totals
 * and is intentionally ignored for a single-period export.
 */
function workbookToNestedDictionary(workbook) {
    const sheets = {};

    for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(worksheet, {
            header: 1,
            defval: null,
            raw: true,
        });

        const headers = rows[0] ?? [];
        const values = rows[1] ?? [];

        if (headers.length === 0) {
            throw new Error(
                `Worksheet "${sheetName}" does not contain headers.`,
            );
        }

        if (values.length === 0) {
            throw new Error(
                `Worksheet "${sheetName}" does not contain report values.`,
            );
        }

        const duplicateHeaders = headers.filter(
            (header, index) => (
                header !== null
                && header !== undefined
                && headers.indexOf(header) !== index
            ),
        );

        if (duplicateHeaders.length > 0) {
            throw new Error(
                `Worksheet "${sheetName}" contains duplicate headers: `
                + `${[...new Set(duplicateHeaders)].join(', ')}`,
            );
        }

        sheets[sheetName] = Object.fromEntries(
            headers
                .map((header, index) => [
                    header,
                    values[index] ?? null,
                ])
                .filter(([header]) => (
                    header !== null
                    && header !== undefined
                    && String(header).trim() !== ''
                )),
        );
    }

    return sheets;
}

let exitCode = 0;
let statusMessage;

try {
    log.info('Reading Actor input.');

    const input = await Actor.getInput();

    const {
        url = 'https://laynes.revelup.com/reports/sales_summary/',
        username,
        password,
        establishment = 'Leander',
        override_flag = false,
        override_startDate,
        override_endDate,
        startTime,
        startMeridiem,
        endTime,
        endMeridiem,
    } = input ?? {};

    let startDate;
    let endDate;

    if (override_flag) {
        if (!override_startDate || !override_endDate) {
            throw new Error(
                'override_startDate and override_endDate are required '
                + 'when override_flag is true.',
            );
        }

        startDate = override_startDate;
        endDate = override_endDate;
    } else {
        const today = new Date();

        startDate = calculateStartDate(today);
        endDate = calculateEndDate(today);
    }

    log.info('Actor input loaded.', {
        hasUsername: Boolean(username),
        hasPassword: Boolean(password),
        establishment,
        override_flag,
        startDate,
        startTime,
        startMeridiem,
        endDate,
        endTime,
        endMeridiem,
    });

    if (!username || !password) {
        throw new Error('Both username and password are required.');
    }

    if (
        !startTime
        || !startMeridiem
        || !endTime
        || !endMeridiem
    ) {
        throw new Error(
            'Start and end times and AM/PM values are required.',
        );
    }

    if (establishment !== 'Leander') {
        throw new Error(
            `The current Actor version only supports Leander. `
            + `Received: ${establishment}`,
        );
    }

    validateDate(startDate, 'startDate');
    validateDate(endDate, 'endDate');
    validateTime(startTime, 'startTime');
    validateTime(endTime, 'endTime');

    const normalizedStartMeridiem =
        startMeridiem.trim().toUpperCase();

    const normalizedEndMeridiem =
        endMeridiem.trim().toUpperCase();

    if (
        normalizedStartMeridiem !== 'AM'
        || normalizedEndMeridiem !== 'AM'
    ) {
        throw new Error(
            'The current Actor version supports AM report times only.',
        );
    }

    const targetEstablishment = 'Leander';

    const crawler = new PlaywrightCrawler({
        maxRequestsPerCrawl: 1,
        maxRequestRetries: 0,
        requestHandlerTimeoutSecs: 240,

        async requestHandler({ page, request }) {
            log.info(`Opening Revel portal: ${request.url}`);

            await page.goto(request.url, {
                waitUntil: 'domcontentloaded',
                timeout: 30_000,
            });

            await saveScreenshot(page, 'REVEL_LOGIN_START');

            const usernameField = page.locator('#username');

            await usernameField.waitFor({
                state: 'visible',
                timeout: 15_000,
            });

            await usernameField.fill(username);

            log.info('Username entered. Clicking Continue.');

            await page
                .getByRole('button', {
                    name: 'Continue',
                    exact: true,
                })
                .click();

            const passwordField = page.locator(
                'input[type="password"]',
            );

            await passwordField.waitFor({
                state: 'visible',
                timeout: 20_000,
            });

            log.info('Password field appeared.');

            await saveScreenshot(page, 'REVEL_PASSWORD_STEP');

            await passwordField.fill(password);

            const loginButton = page
                .locator(
                    'button[type="submit"]:visible, '
                    + 'input[type="submit"]:visible',
                )
                .last();

            await loginButton.waitFor({
                state: 'visible',
                timeout: 15_000,
            });

            const buttonText =
                (await loginButton.textContent())?.trim()
                || (await loginButton.getAttribute('value'))
                || 'Submit';

            log.info(`Clicking final login button: ${buttonText}`);

            await loginButton.click();

            await passwordField.waitFor({
                state: 'hidden',
                timeout: 30_000,
            });

            await page.waitForLoadState('domcontentloaded');

            log.info(`Login completed. Current URL: ${page.url()}`);

            if (!page.url().includes('/reports/sales_summary')) {
                log.info(`Navigating to report URL: ${url}`);

                await page.goto(url, {
                    waitUntil: 'domcontentloaded',
                    timeout: 30_000,
                });
            }

            const establishmentText = page.locator(
                '[data-cy="header-establishment-text"]',
            );

            await establishmentText.waitFor({
                state: 'visible',
                timeout: 20_000,
            });

            const currentEstablishment =
                (await establishmentText.textContent())?.trim()
                || 'Unknown';

            log.info(
                `Current establishment before selection: `
                + `${currentEstablishment}`,
            );

            if (currentEstablishment !== targetEstablishment) {
                await establishmentText.click();

                log.info('Clicked the establishment name.');

                const leanderOption = page
                    .locator('span.fancytree-title')
                    .filter({
                        hasText: '42 | Leander',
                    });

                await leanderOption.waitFor({
                    state: 'visible',
                    timeout: 30_000,
                });

                log.info('Establishment panel opened.');

                await saveScreenshot(
                    page,
                    'REVEL_ESTABLISHMENT_LIST',
                );

                log.info(
                    `Selecting establishment: ${targetEstablishment}`,
                );

                await leanderOption.click();
            }

            const selectedHeader = page
                .locator('[data-cy="header-establishment-text"]')
                .filter({
                    hasText: /^\s*Leander\s*$/,
                });

            await selectedHeader.waitFor({
                state: 'visible',
                timeout: 30_000,
            });

            const selectedEstablishment =
                (await selectedHeader.textContent())?.trim();

            if (selectedEstablishment !== targetEstablishment) {
                throw new Error(
                    `Expected establishment "${targetEstablishment}", `
                    + `but found "${selectedEstablishment}".`,
                );
            }

            log.info(
                `Establishment selected successfully: `
                + `${selectedEstablishment}`,
            );

            const salesSummaryHeading = page.getByRole('heading', {
                name: 'Sales Summary',
                exact: true,
            });

            await salesSummaryHeading.waitFor({
                state: 'visible',
                timeout: 30_000,
            });

            await saveScreenshot(
                page,
                'REVEL_LEANDER_SALES_SUMMARY',
            );

            await applyReportFilters(page);

            const dateRangeDropdown = page.locator(
                '.report-date-row .ico-f-to-down',
            );

            await dateRangeDropdown.waitFor({
                state: 'visible',
                timeout: 20_000,
            });

            log.info('Opening the Sales Summary date-range dropdown.');

            await dateRangeDropdown.click();

            const visibleDatePicker = page.locator(
                '.daterangepicker:visible',
            );

            await visibleDatePicker.waitFor({
                state: 'visible',
                timeout: 20_000,
            });

            await saveScreenshot(
                page,
                'REVEL_DATE_RANGE_OPEN',
            );

            log.info(
                `Setting internal report range: `
                + `${startDate} ${startTime} `
                + `${normalizedStartMeridiem} through `
                + `${endDate} ${endTime} `
                + `${normalizedEndMeridiem}.`,
            );

            const pickerResult = await page.evaluate(
                ({
                    startDateValue,
                    startTimeValue,
                    startMeridiemValue,
                    endDateValue,
                    endTimeValue,
                    endMeridiemValue,
                }) => {
                    const $ = window.jQuery;
                    const moment = window.moment;

                    if (!$) {
                        throw new Error(
                            'jQuery is not available on the page.',
                        );
                    }

                    if (!moment) {
                        throw new Error(
                            'Moment.js is not available on the page.',
                        );
                    }

                    const candidates = $('*').filter(
                        function findPicker() {
                            return Boolean(
                                $(this).data('daterangepicker'),
                            );
                        },
                    );

                    if (candidates.length === 0) {
                        throw new Error(
                            'Unable to locate Revel '
                            + 'daterangepicker instance.',
                        );
                    }

                    let picker = null;

                    candidates.each(
                        function selectVisiblePicker() {
                            const candidate =
                                $(this).data('daterangepicker');

                            if (
                                !picker
                                && candidate?.container
                                && candidate.container.is(':visible')
                            ) {
                                picker = candidate;
                            }
                        },
                    );

                    if (!picker) {
                        picker = $(candidates[0])
                            .data('daterangepicker');
                    }

                    const startDateTime = moment(
                        `${startDateValue} `
                        + `${startTimeValue} `
                        + `${startMeridiemValue}`,
                        'MM/DD/YYYY hh:mm A',
                        true,
                    );

                    const endDateTime = moment(
                        `${endDateValue} `
                        + `${endTimeValue} `
                        + `${endMeridiemValue}`,
                        'MM/DD/YYYY hh:mm A',
                        true,
                    );

                    if (!startDateTime.isValid()) {
                        throw new Error(
                            'The requested start date/time is invalid.',
                        );
                    }

                    if (!endDateTime.isValid()) {
                        throw new Error(
                            'The requested end date/time is invalid.',
                        );
                    }

                    if (endDateTime.isBefore(startDateTime)) {
                        throw new Error(
                            'The report end date/time cannot be '
                            + 'before the start date/time.',
                        );
                    }

                    if (
                        typeof picker.setStartDate !== 'function'
                        || typeof picker.setEndDate !== 'function'
                    ) {
                        throw new Error(
                            'The Revel daterangepicker does not '
                            + 'expose its date-setting methods.',
                        );
                    }

                    picker.setStartDate(startDateTime);
                    picker.setEndDate(endDateTime);

                    if (typeof picker.updateView === 'function') {
                        picker.updateView();
                    }

                    if (
                        typeof picker.updateCalendars === 'function'
                    ) {
                        picker.updateCalendars();
                    }

                    if (
                        typeof picker.updateFormInputs === 'function'
                    ) {
                        picker.updateFormInputs();
                    }

                    return {
                        startDate: picker.startDate.format(
                            'MM/DD/YYYY hh:mm A',
                        ),
                        endDate: picker.endDate.format(
                            'MM/DD/YYYY hh:mm A',
                        ),
                        hasClickApply:
                            typeof picker.clickApply === 'function',
                    };
                },
                {
                    startDateValue: startDate,
                    startTimeValue: startTime,
                    startMeridiemValue:
                        normalizedStartMeridiem,
                    endDateValue: endDate,
                    endTimeValue: endTime,
                    endMeridiemValue:
                        normalizedEndMeridiem,
                },
            );

            log.info(
                `Internal picker range: `
                + `${pickerResult.startDate} through `
                + `${pickerResult.endDate}`,
            );

            if (!pickerResult.hasClickApply) {
                throw new Error(
                    'The Revel daterangepicker does not expose '
                    + 'its Apply method.',
                );
            }

            await saveScreenshot(
                page,
                'REVEL_DATE_RANGE_POPULATED',
            );

            log.info(
                'Applying the date range through the picker API.',
            );

            await page.evaluate(() => {
                const $ = window.jQuery;

                if (!$) {
                    throw new Error(
                        'jQuery is not available during Apply.',
                    );
                }

                const candidates = $('*').filter(
                    function findPicker() {
                        return Boolean(
                            $(this).data('daterangepicker'),
                        );
                    },
                );

                let picker = null;

                candidates.each(
                    function selectVisiblePicker() {
                        const candidate =
                            $(this).data('daterangepicker');

                        if (
                            !picker
                            && candidate?.container
                            && candidate.container.is(':visible')
                        ) {
                            picker = candidate;
                        }
                    },
                );

                if (!picker && candidates.length > 0) {
                    picker = $(candidates[0])
                        .data('daterangepicker');
                }

                if (!picker) {
                    throw new Error(
                        'Unable to locate the daterangepicker '
                        + 'during Apply.',
                    );
                }

                if (typeof picker.clickApply !== 'function') {
                    throw new Error(
                        'The Revel daterangepicker does not '
                        + 'expose clickApply().',
                    );
                }

                picker.clickApply();
            });

            await visibleDatePicker.waitFor({
                state: 'hidden',
                timeout: 30_000,
            });

            log.info(
                'Date picker closed. Waiting for Revel to refresh '
                + 'the report.',
            );

            /*
             * Do not rely on a fixed delay here. Revel updates the date
             * label before/while it asynchronously rebuilds the report,
             * and slower requests can otherwise export the prior period.
             */
            await page.waitForFunction(
                ({
                    expectedStartDate,
                    expectedEndDate,
                }) => {
                    const normalizeDate = (value) => {
                        const match = String(value).match(
                            /(\d{1,2})\/(\d{1,2})\/(\d{4})/,
                        );

                        if (!match) return null;

                        const [, month, day, year] = match;

                        return `${month.padStart(2, '0')}/`
                            + `${day.padStart(2, '0')}/${year}`;
                    };

                    const reportDateRow = document.querySelector(
                        '.report-date-row',
                    );

                    if (!reportDateRow) return false;

                    const displayedDates = (
                        reportDateRow.textContent ?? ''
                    )
                        .match(/\d{1,2}\/\d{1,2}\/\d{4}/g)
                        ?.map(normalizeDate);

                    if (!displayedDates || displayedDates.length < 2) {
                        return false;
                    }

                    return (
                        displayedDates[0]
                            === normalizeDate(expectedStartDate)
                        && displayedDates[1]
                            === normalizeDate(expectedEndDate)
                    );
                },
                {
                    expectedStartDate: startDate,
                    expectedEndDate: endDate,
                },
                {
                    timeout: 90_000,
                    polling: 500,
                },
            );

            const displayedRange = (
                await page.locator('.report-date-row').innerText()
            ).replace(/\s+/g, ' ').trim();

            log.info(
                `Revel displays the requested report range: `
                + `${displayedRange}`,
            );

            await waitForReportIdle(page);

            await saveScreenshot(
                page,
                'REVEL_SALES_SUMMARY_APPLIED',
            );

            log.info(
                'Sales Summary refreshed for the requested date range.',
            );

            /*
             * Open the three-dot export menu.
             */
            const exportMenuButton = page
                .locator('.header-more .button-more:visible')
                .first();

            await exportMenuButton.waitFor({
                state: 'visible',
                timeout: 20_000,
            });

            log.info('Opening the report export menu.');

            await exportMenuButton.click();

            /*
             * Select the Excel export option.
             */
            const excelExportLink = page
                .locator('[data-exporttype="excel"]:visible')
                .first();

            await excelExportLink.waitFor({
                state: 'visible',
                timeout: 20_000,
            });

            await saveScreenshot(
                page,
                'REVEL_EXPORT_MENU_OPEN',
            );

            log.info('Downloading the Sales Summary Excel report.');

            const downloadPromise = page.waitForEvent(
                'download',
                {
                    timeout: 60_000,
                },
            );

            await excelExportLink.click();

            const download = await downloadPromise;

            const downloadFailure = await download.failure();

            if (downloadFailure) {
                throw new Error(
                    `Excel download failed: ${downloadFailure}`,
                );
            }

            const temporaryFilePath = await download.path();

            if (!temporaryFilePath) {
                throw new Error(
                    'Playwright did not provide a path '
                    + 'for the downloaded file.',
                );
            }

            const excelBuffer = await readFile(
                temporaryFilePath,
            );

            /*
             * Parse the downloaded Excel workbook directly from
             * memory and map every sheet's row 1 headers to its
             * row 2 values.
             */
            const workbook = XLSX.read(excelBuffer, {
                type: 'buffer',
                cellDates: false,
            });

            const sheets = workbookToNestedDictionary(workbook);

            const fieldCountBySheet = Object.fromEntries(
                Object.entries(sheets).map(
                    ([sheetName, fields]) => [
                        sheetName,
                        Object.keys(fields).length,
                    ],
                ),
            );

            const allRevenueCenters =
                sheets['All revenue centers'];

            if (!allRevenueCenters) {
                throw new Error(
                    'The "All revenue centers" sheet was not found.',
                );
            }

            const rawNetSales = allRevenueCenters['Net Sales'];

            if (
                rawNetSales === undefined
                || rawNetSales === null
                || rawNetSales === ''
            ) {
                throw new Error(
                    'The "Net Sales" field was not found in '
                    + 'the "All revenue centers" sheet.',
                );
            }

            /*
            * XLSX normally returns this as a number. This fallback also
            * handles values formatted like "$5,350.50".
            */
            const netSales = typeof rawNetSales === 'number'
                ? rawNetSales
                : Number(
                    String(rawNetSales).replace(/[$,\s]/g, ''),
                );

            if (!Number.isFinite(netSales)) {
                throw new Error(
                    `Net Sales is not numeric: ${rawNetSales}`,
                );
            }

            /*
            * Convert MM/DD/YYYY into YYYY-MM-DD.
            * Example: 07/04/2026 becomes 2026-07-04.
            */
            const [startMonth, startDay, startYear] =
                startDate.split('/');

            const formattedStartDate =
                `${startYear}-${startMonth.padStart(2, '0')}`
                + `-${startDay.padStart(2, '0')}`;

            const dailySalesRow = {
                id: `${formattedStartDate}_${selectedEstablishment}`,
                location: selectedEstablishment,
                business_date: formattedStartDate,

                time_from: allRevenueCenters['Time From'],
                time_to: allRevenueCenters['Time To'],

                taxable_sales: allRevenueCenters['Taxable Sales'],
                total_product_sales:
                    allRevenueCenters['Total Product Sales'],
                gross_sales: allRevenueCenters['Gross Sales'],
                net_sales: allRevenueCenters['Net Sales'],
                total_payments: allRevenueCenters['Total Payments'],
                net_account_for: allRevenueCenters['Net Account For'],

                to_go: allRevenueCenters['To Go'],
                to_go_percent: allRevenueCenters['To Go P'],
                eat_in: allRevenueCenters['Eat In'],
                eat_in_percent: allRevenueCenters['Eat In P'],
                drive_through: allRevenueCenters['Drive Through'],
                drive_through_percent:
                    allRevenueCenters['Drive Through P'],
                pickup: allRevenueCenters['Pickup'],
                pickup_percent: allRevenueCenters['Pickup P'],
                dd_marketplace: allRevenueCenters['DD Marketplace'],
                dd_marketplace_percent:
                    allRevenueCenters['DD Marketplace P'],
                uber_eats: allRevenueCenters['Uber Eats'],
                uber_eats_percent: allRevenueCenters['Uber Eats P'],

                item_discounts: allRevenueCenters['Item Discounts'],
                order_discounts_total:
                    allRevenueCenters['Order Discounts Total'],
                total_discounts: allRevenueCenters['Total Discounts'],
                sales_tax: allRevenueCenters['Sales Tax'],
                total_tax_surcharge_service:
                    allRevenueCenters['Total Tax Surcharge Service'],

                cash_total: allRevenueCenters['Cash Total'],
                credit_total: allRevenueCenters['Credit Total'],
                visa_total: allRevenueCenters['Visa Total'],
                mastercard_total: allRevenueCenters['Mastercard Total'],
                american_express_total:
                    allRevenueCenters['American Express Total'],
                discover_total: allRevenueCenters['Discover Total'],
                custom_payment_total:
                    allRevenueCenters['Custom Payment Total'],
                uber_eats_total: allRevenueCenters['Uber Eats Total'],
                dd_marketplace_total:
                    allRevenueCenters['Dd Marketplace Total'],
                actual_total_cash_to_business:
                    allRevenueCenters['Actual Total Cash To Business'],
                payments_captured_amount:
                    allRevenueCenters['Payments Captured Amount'],

                refunds_total: allRevenueCenters['Refunds Total'],
                refunds_count: allRevenueCenters['Refunds Count'],
                voided_total: allRevenueCenters['Voided Total'],
                voided_items: allRevenueCenters['Voided Items'],

                adj_non_cash_tips_total_minus_cash:
                    allRevenueCenters[
                        'Adj Non Cash Tips Total Minus Cash'
                    ],
                total_orders: allRevenueCenters['Total Orders'],
                avg_sale: allRevenueCenters['Avg Sale'],
                cash_due_house: allRevenueCenters['Cash Due House'],
                cash_due_payments:
                    allRevenueCenters['Cash Due Payments'],
                cash_payments: allRevenueCenters['Cash Payments'],

                counter_sales_net_sales:
                    allRevenueCenters['Counter Sales Net Sales'],
                counter_sales_net_sales_percent:
                    allRevenueCenters[
                        'Counter Sales Net Sales Percent'
                    ],
                drive_thru_sales_net_sales:
                    allRevenueCenters['Drive Thru Sales Net Sales'],
                drive_thru_sales_net_sales_percent:
                    allRevenueCenters[
                        'Drive Thru Sales Net Sales Percent'
                    ],

                // Preserve every field from the source report.
                raw_data: allRevenueCenters,

                extracted_at: new Date().toISOString(),
            };

            const {
                data: savedSupabaseRows,
                error: supabaseError,
            } = await supabase
                .from('daily-sales-summary-all-revenue-operations')
                .upsert(
                    dailySalesRow,
                    {
                        onConflict: 'id',
                    },
                )
                .select();

            if (supabaseError) {
                throw new Error(
                    `Unable to write daily sales to Supabase: `
                    + `${supabaseError.message}`,
                );
            }

            log.info(
                `Supabase row saved: ${dailySalesRow.id}`,
                {
                    location: dailySalesRow.Location,
                    netSales: dailySalesRow['Net Sales'],
                },
            );

            await Actor.pushData({
                status: 'success',
                portalUrl: page.url(),
                pageTitle: await page.title(),
                establishment: selectedEstablishment,
                report: 'Sales Summary',
                startDate,
                startTime,
                startMeridiem: normalizedStartMeridiem,
                endDate,
                endTime,
                endMeridiem: normalizedEndMeridiem,
                sheetNames: workbook.SheetNames,
                fieldCountBySheet,
                // sourceFilename: download.suggestedFilename(),
                sourceSizeBytes: excelBuffer.length,
                timestamp: new Date().toISOString(),
                message:
                    'Logged into Revel, selected Leander, '
                    + 'applied the Sales Summary date range, '
                    + 'parsed every worksheet, and saved the daily sales row.',
            });
        },

        async failedRequestHandler({ page, request }, error) {
            log.error(`Revel extraction failed: ${error.message}`);

            if (page) {
                try {
                    await saveScreenshot(
                        page,
                        'REVEL_EXTRACTION_FAILURE',
                    );
                } catch (screenshotError) {
                    log.warning(
                        `Unable to save failure screenshot: `
                        + `${screenshotError.message}`,
                    );
                }
            }

            await Actor.pushData({
                status: 'failed',
                portalUrl: page
                    ? page.url()
                    : request.url,
                pageTitle: page
                    ? await page.title().catch(() => '')
                    : '',
                establishment: targetEstablishment,
                report: 'Sales Summary',
                startDate,
                startTime,
                startMeridiem: normalizedStartMeridiem,
                endDate,
                endTime,
                endMeridiem: normalizedEndMeridiem,
                timestamp: new Date().toISOString(),
                message: error.message,
            });
        },
    });

    await crawler.run([url]);
} catch (error) {
    const failure = error instanceof Error
        ? error
        : new Error(String(error));

    exitCode = 1;
    statusMessage = failure.message;
    log.exception(failure, failure.message);
} finally {
    await Actor.exit({
        exitCode,
        ...(statusMessage ? { statusMessage } : {}),
    });
}