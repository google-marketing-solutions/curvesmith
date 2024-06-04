/**
 * @license
 * Copyright 2024 Google LLC.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *     https://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */

/**
 * @fileoverview The main entry point for the CurveSmith solution.
 */

import * as am_handler from './ad_manager_handler';
import {TaskProgress} from './assets/common';
import {DialogSettings, LineItemPreview} from './assets/upload';
import {FlightDetails} from './custom_curve';
import {LineItemRow, SheetHandler, SpreadsheetHandler} from './sheet_handler';
import * as ad_manager from './typings/ad_manager';

/**
 * A map of callback functions that can be called from client side JavaScript.
 */
const CALLBACK_FUNCTIONS: {[id: string]: (...args: any[]) => any} = {
  /* Server Features */
  'applyHistorical': applyHistorical,
  'copyTemplate': copyTemplate,
  'loadLineItems': loadLineItems,
  'showPreviewDialog': showPreviewDialog,
  'showUploadDialog': showUploadDialog,
  'uploadLineItems': uploadLineItems,

  /* Sidebar and Dialog communication */
  'clearTaskProgress': clearTaskProgress,
  'getTaskProgress': getTaskProgress,

  /* Miscellaneous */
  'initializeSpreadsheet': initializeSpreadsheet,
};

/**
 * Calls the specified callback function with the provided arguments.
 * @throws An error if the function is not a callable.
 */
export function callback(functionName: string, args: object[]): unknown {
  const func = CALLBACK_FUNCTIONS[functionName];

  if (func) {
    return func.apply(func, args);
  } else {
    throw new Error(`${functionName} is not callable`);
  }
}

/**
 * When the spreadsheet is first opened, add a solution specific menu that will
 * facilitate the creation and upload of custom curves.
 */
export function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Custom Curves')
    .addItem('Load Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * Returns the content of an HTML component file (i.e. CSS or JavaScript) to be
 * used in conjunction with `HtmlTemplate.evaluate`. A client-side HTML file can
 * call this function with a scriptlet (i.e. `<?!= include('common.html') ?>`).
 */
export function include(filename: string): string {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Restores the sidebar UI if it has been closed. Otherwise, reloads the sidebar
 * back to the default state.
 */
export function showSidebar(): void {
  clearTaskProgress();

  const fileName =
    'dist/' +
    'sidebar';

  const userInterface = HtmlService.createTemplateFromFile(fileName)
    .evaluate()
    .setTitle('Custom Curves');

  SpreadsheetApp.getUi().showSidebar(userInterface);
}

/**
 * Applies historical delivery pacing to line items referenced within the active
 * sheet. Selection behavior follows:
 * - If no line items are selected, then all line items will be affected
 * - If line items are selected, then only those line items will be affected
 */
function applyHistorical(
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
  sheetHandler: SheetHandler = getSheetHandler(),
): void {
  const lineItemRows = sheetHandler.getSelectedLineItems();

  if (lineItemRows.length === 0) {
    throw new Error('No line items are selected');
  }

  const WARNING_MESSAGE =
    'Setting the delivery pacing source to "historical" cannot be reverted. ' +
    'Any existing custom curves will be lost. Are you sure you want to proceed?';

  if (shouldCancelAction(WARNING_MESSAGE)) {
    throw new Error('User cancelled application of historical delivery pacing');
  }

  const lineItemIds = lineItemRows.map((lineItemRow) => lineItemRow.id);
  const lineItems = getLineItemsWithIds(lineItemIds, adManagerHandler);

  adManagerHandler.applyHistoricalToLineItems(lineItems);
  adManagerHandler.uploadLineItems(lineItems);
}

/**
 * Creates a new copy of the custom curve template sheet. A user may also just
 * directly duplicate an existing template within the Google Sheets interface.
 */
function copyTemplate(
  spreadsheetHandler: SpreadsheetHandler = getSpreadsheetHandler(),
): void {
  spreadsheetHandler.copyTemplate();
}

/** Clears the current and total progress values for the active task. */
function clearTaskProgress(): void {
  const userProperties = PropertiesService.getUserProperties();

  userProperties.deleteProperty('current');
  userProperties.deleteProperty('total');
}

/**
 * Returns an `AdManagerHandler` based on network configuration details stored
 * within the active spreadsheet.
 */
function getAdManagerHandler(
  spreadsheetHandler: SpreadsheetHandler = getSpreadsheetHandler(),
): am_handler.AdManagerHandler {
  const networkId: string = spreadsheetHandler.getNetworkId();
  const apiVersion: string = spreadsheetHandler.getApiVersion();

  const client = am_handler.createAdManagerClient(networkId, apiVersion);

  return new am_handler.AdManagerHandler(client);
}

/**
 * Returns a line item filter. The minimum bounds for the flight window are
 * derived from the scheduled events in order to ensure that line items can
 * support the full curve.
 */
function getLineItemFilter(
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
  sheetHandler: SheetHandler = getSheetHandler(),
): am_handler.LineItemFilter {
  const adUnitId = sheetHandler.getAdUnitId();
  const events = sheetHandler.getScheduledEvents();

  if (events.length === 0) {
    throw new Error('No scheduled events are specified');
  }

  // Scheduled events must be sorted in ascending order, or an error will be
  // thrown. Consequently, the first event will always be the earliest.
  const latestStartDate = events[0].start;

  // The last event will always be the most recent, however it may cover a date
  // range that has already passed. Line items that have already ended are not
  // eligible, so we ensure that the end date is in the future.
  const earliestEndDate = new Date(
    Math.max(
      events[events.length - 1].end.getTime(),
      new Date().setHours(23, 59, 59, 999),
    ),
  );

  return {
    adUnitIds: adManagerHandler.getAdUnitIds(adUnitId),
    latestStartDate,
    earliestEndDate,
  };
}

/**
 * Returns a list of selected line items and their corresponding curve previews.
 */
function getLineItemPreviews(
  sheetHandler: SheetHandler = getSheetHandler(),
): LineItemPreview[] {
  const curveTemplate = sheetHandler.getCurveTemplate();
  const lineItemRows = sheetHandler.getSelectedLineItems();

  if (lineItemRows.length === 0) {
    throw new Error('No line items are selected');
  }

  return lineItemRows.map((lineItemRow) => {
    const flight = new FlightDetails(
      lineItemRow.startDate,
      lineItemRow.endDate,
      lineItemRow.impressionGoal,
    );

    return {
      id: lineItemRow.id,
      name: lineItemRow.name,
      startDate: flight.start.toISOString(),
      endDate: flight.end.toISOString(),
      impressionGoal: flight.impressionGoal,
      curveGoals: curveTemplate
        .generateCurveSegments(flight)
        .map((segment) => ({
          description: segment.description,
          startDate: segment.start.toISOString(),
          goalPercent: segment.goalPercent,
          impressionGoal: flight.impressionGoal * (segment.goalPercent / 100),
        })),
    };
  });
}

function getLineItemsWithIds(
  lineItemIds: number[],
  adManagerHandler: am_handler.AdManagerHandler,
): ad_manager.LineItem[] {
  if (lineItemIds.length === 0) {
    throw new Error('No line items are selected');
  }

  const lineItems: ad_manager.LineItem[] = [];

  let offset = 0;
  let lineItemPage: ad_manager.LineItemPage;

  do {
    lineItemPage = adManagerHandler.getLineItemsWithIds(lineItemIds, offset);

    lineItems.push(...lineItemPage.results);

    offset += am_handler.AdManagerHandler.AD_MANAGER_API_PAGE_LIMIT;

    setTaskProgress(offset, lineItemPage.totalResultSetSize);
  } while (offset < lineItemPage.totalResultSetSize);

  return lineItems;
}

/** Returns a `SheetHandler` based on the active sheet. */
function getSheetHandler(): SheetHandler {
  return new SheetHandler(SpreadsheetApp.getActiveSheet());
}

/** Returns a `SpreadsheetHandler` based on the active spreadsheet. */
function getSpreadsheetHandler(): SpreadsheetHandler {
  return new SpreadsheetHandler(SpreadsheetApp.getActiveSpreadsheet());
}

/** Returns the current and total progress values for the active task. */
function getTaskProgress(): TaskProgress {
  const userProperties = PropertiesService.getUserProperties();

  const current = userProperties.getProperty('current');
  const total = userProperties.getProperty('total');

  return {
    current: current ? Number(current) : 0,
    total: total ? Number(total) : 0,
  };
}

/**
 * Initializes the active spreadsheet for use.
 *
 * Currently this entails:
 * - setting the spreadsheet time zone to match the Ad Manager network time zone
 */
function initializeSpreadsheet(): void {
  const spreadsheetHandler = getSpreadsheetHandler();
  const adManagerHandler = getAdManagerHandler(spreadsheetHandler);

  spreadsheetHandler.updateSpreadsheetTimeZone(
    adManagerHandler.getTimeZoneId(),
  );
}

/**
 * Retrieves line items from Ad Manager and writes pertitent metadata to the
 * active sheet. We limit line items to only those that meet the requirements
 * for custom delivery curves, including flight date range and ad unit ID.
 */
function loadLineItems(
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
  sheetHandler: SheetHandler = getSheetHandler(),
): void {
  const WARNING_MESSAGE =
    'Loading line items will overwrite line item data in the active sheet. ' +
    'All other data will be unaffected. Are you sure you want to proceed?';

  if (shouldCancelAction(WARNING_MESSAGE)) {
    throw new Error('User cancelled loading line items');
  }

  const filter = getLineItemFilter(adManagerHandler, sheetHandler);

  sheetHandler.clearLineItems();

  let offset = 0;
  let lineItemPage: ad_manager.LineItemPage;

  do {
    lineItemPage = adManagerHandler.getLineItemsByFilter(filter, offset);

    const lineItemRows: LineItemRow[] = [];

    for (const lineItem of lineItemPage.results) {
      lineItemRows.push({
        selected: false,
        id: lineItem.id,
        name: lineItem.name,
        startDate: adManagerHandler.getDateString(lineItem.startDateTime),
        endDate: adManagerHandler.getDateString(lineItem.endDateTime),
        impressionGoal: lineItem.primaryGoal.units,
      });
    }

    sheetHandler.appendLineItems(lineItemRows);

    offset += am_handler.AdManagerHandler.AD_MANAGER_API_PAGE_LIMIT;

    setTaskProgress(offset, lineItemPage.totalResultSetSize);
  } while (offset < lineItemPage.totalResultSetSize);
}

/** Sets the current and total progress values for the active task. */
function setTaskProgress(current: number, total: number): void {
  const userProperties = PropertiesService.getUserProperties();

  if (current > total) {
    current = total; // Ensure current is never greater than total
  }

  userProperties.setProperty('current', current.toString());
  userProperties.setProperty('total', total.toString());
}

/**
 * Displays a confirmation dialog to the user allowing them to cancel the
 * action.
 * @param message The message to display in the dialog.
 * @return True if the user cancels the action, false otherwise.
 */
function shouldCancelAction(message: string): boolean {
  const UI = SpreadsheetApp.getUi();

  return UI.alert(message, UI.ButtonSet.YES_NO) === UI.Button.NO;
}

/**
 * Displays a modal dialog to the user. In order to avoid extra server-side
 * calls, required parameters are passed to the dialog via scriplets.
 */
function showDialog(title: string, previewOnly: boolean): void {
  const lineItemPreviews = getLineItemPreviews();

  const fileName =
    'dist/' +
    'upload';

  const template = HtmlService.createTemplateFromFile(fileName);

  const spreadsheetHandler = getSpreadsheetHandler();

  const dialogSettings: DialogSettings = {
    networkId: spreadsheetHandler.getNetworkId(),
    showApproveAll: spreadsheetHandler.getShowApproveAll(),
    previewOnly,
  };

  template['dialogSettings'] = JSON.stringify(dialogSettings);
  template['lineItemPreviews'] = JSON.stringify(lineItemPreviews);

  const userInterface = template.evaluate().setWidth(500).setHeight(500);

  SpreadsheetApp.getUi().showModalDialog(userInterface, title);
}

/** Displays a dialog to preview custom curves for selected line items. */
function showPreviewDialog(): void {
  showDialog(/* title= */ 'Curve Preview', /* previewOnly= */ true);
}

/** Displays a dialog to upload custom curves to selected line items. */
function showUploadDialog(): void {
  showDialog(/* title= */ 'Curve Upload', /* previewOnly= */ false);
}

/**
 * Updates the line items with the custom delivery curve specified in the
 * active sheet.
 * @return An array containing status for each line item update requested
 */
function uploadLineItems(
  lineItemIds: number[],
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
  sheetHandler: SheetHandler = getSheetHandler(),
): void {
  const curveTemplate = sheetHandler.getCurveTemplate();

  const lineItems = getLineItemsWithIds(lineItemIds, adManagerHandler);

  adManagerHandler.applyCurveToLineItems(lineItems, curveTemplate);
  adManagerHandler.uploadLineItems(lineItems);
}

global.onOpen = onOpen;
global.callback = callback;
global.include = include;
global.showSidebar = showSidebar;

export const TEST_ONLY = {
  copyTemplate,
  getAdManagerHandler,
  getLineItemPreviews,
  getTaskProgress,
  loadLineItems,
};
