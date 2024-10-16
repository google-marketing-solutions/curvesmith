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

import {AdManagerServerFault} from 'gam_apps_script/ad_manager_error';

/**
 * A map of callback functions that can be called from client side JavaScript.
 */
const CALLBACK_FUNCTIONS: {[id: string]: (...args: any[]) => any} = {
  /* Server Features */
  'applyHistorical': applyHistorical,
  'beginApplyHistorical': beginApplyHistorical,
  'beginLoadLineItems': beginLoadLineItems,
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

const SCRIPT_VERSION = '0.0.3';

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

/** Handles the `onEdit` event for the active sheet. */
export function onEdit(event: GoogleAppsScript.Events.SheetsOnEdit): void {
  const sheetHandler = getSheetHandler();

  sheetHandler.handleEdit(event);
}

/**
 * When the spreadsheet is first opened, add a solution specific menu that will
 * facilitate the creation and upload of custom curves.
 */
export function onOpen(): void {
  SpreadsheetApp.getUi()
    .createMenu('Custom Curves [' + SCRIPT_VERSION + ']')
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
 * The number of objects to request at a time. This value was empirically
 * determined to be the optimal tradeoff between UX and performance.
 */
const AD_MANAGER_API_SUGGESTED_PAGE_SIZE = 50;

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

  userProperties.deleteProperty('action');
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
  const timeZoneId: string = spreadsheetHandler.getSpreadsheetTimeZone();

  const client = am_handler.createAdManagerClient(networkId, apiVersion);

  return new am_handler.AdManagerHandler(client, timeZoneId);
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
  const nameFilter = sheetHandler.getNameFilter().trim();

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
    nameFilter,
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
    lineItemPage = adManagerHandler.getLineItemsWithIds(
      lineItemIds,
      offset,
      AD_MANAGER_API_SUGGESTED_PAGE_SIZE,
    );

    lineItems.push(...lineItemPage.results);

    offset += AD_MANAGER_API_SUGGESTED_PAGE_SIZE;
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

  const action = userProperties.getProperty('action');
  const current = userProperties.getProperty('current');
  const total = userProperties.getProperty('total');

  return {
    action: action ? action : 'Progress',
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
 * Offers the user the option of loading line items from the Ad Manager API
 * and overwriting the active sheet. We limit line items to only those lines
 * that meet the requirements for custom delivery curves, including flight date
 * range, name filter, and ad unit ID.
 *
 * Parsing SOAP objects returned by the Ad Manager API is prohibitively slow, so
 * line items are requested concurrently and processed in batches. This function
 * returns the total number of line items that match the current filter
 * criteria, however, the actual line items requests are kicked off in parallel
 * by the sidebar UI.
 */
function beginLoadLineItems(
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
  sheetHandler: SheetHandler = getSheetHandler(),
): number {
  const WARNING_MESSAGE =
    'Loading line items will overwrite line item data in the active sheet. ' +
    'All other data will be unaffected. Are you sure you want to proceed?';

  if (shouldCancelAction(WARNING_MESSAGE)) {
    throw new Error('User cancelled loading line items');
  }

  sheetHandler.clearLineItems();

  const filter = getLineItemFilter(adManagerHandler, sheetHandler);

  const lineItemCount = adManagerHandler.getLineItemCount(filter);

  if (lineItemCount === 0) {
    throw new Error('No suitable line items found');
  }

  setTaskProgress('Retrieved', 0, lineItemCount);

  return lineItemCount;
}

/**
 * Offers the user the option of applying historical delivery pacing to the
 * selected line items. If the user chooses to proceed, then a list of line item
 * IDs is returned for subsequent processing.
 */
function beginApplyHistorical(
  sheetHandler: SheetHandler = getSheetHandler(),
): number[] {
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

  return lineItemRows.map((lineItemRow) => lineItemRow.id);
}

/**
 * Retrieves line items from the Ad Manager API and caches them in the script
 * properties cache. Once all expected line items have been cached, the line
 * item metadata is written to the active sheet.
 * @param offset The offset to use for the request
 * @param limit The number of line items to request
 */
function loadLineItems(
  offset: number,
  limit: number,
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
  sheetHandler: SheetHandler = getSheetHandler(),
): void {
  const filter = getLineItemFilter(adManagerHandler, sheetHandler);

  const lineItemDtoPage = adManagerHandler.getLineItemDtoPage(
    filter,
    offset,
    limit,
  );

  const scriptProperties = PropertiesService.getScriptProperties();

  scriptProperties.setProperty(
    'lineItemDtoPage' + offset,
    JSON.stringify(lineItemDtoPage.values),
  );

  const keys = scriptProperties.getKeys();

  const validKeys = keys.filter((key) => key.startsWith('lineItemDtoPage'));

  const taskProgress = getTaskProgress();

  setTaskProgress('Retrieved', validKeys.length * limit, taskProgress.total);

  if (validKeys.length === Math.ceil(taskProgress.total / limit)) {
    writeLineItemMetadata(validKeys, sheetHandler);
  }
}

/**
 * Writes line item metadata to the active sheet.
 * @param cacheKeys The keys of the script properties to read
 */
function writeLineItemMetadata(
  cacheKeys: string[],
  sheetHandler: SheetHandler = getSheetHandler(),
): void {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lineItemRows: LineItemRow[] = [];

  for (const key of cacheKeys) {
    const serializedDtos = scriptProperties.getProperty(key) ?? '';
    const lineItemDtos = JSON.parse(serializedDtos) as am_handler.LineItemDto[];

    for (const lineItemDto of lineItemDtos) {
      lineItemRows.push({
        selected: false,
        id: lineItemDto.id,
        name: lineItemDto.name,
        startDate: lineItemDto.startDate,
        endDate: lineItemDto.endDate,
        impressionGoal: lineItemDto.impressionGoal,
      });
    }

    scriptProperties.deleteProperty(key);
  }

  sheetHandler.writeLineItems(lineItemRows);
}

/** Sets the current and total progress values for the active task. */
function setTaskProgress(action: string, current: number, total: number): void {
  const userProperties = PropertiesService.getUserProperties();

  if (current > total) {
    current = total; // Ensure current is never greater than total
  }

  userProperties.setProperty('action', action);
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

/** Holds information about a line item that failed to upload. */
interface LineItemError {
  lineItemId: number;
  errorString: string;
}

/**
 * Returns a list of line item errors based on the Ad Manager API response.
 * @param serverFault The error returned from the Ad Manager API
 * @param lineItems The line items that were sent to the Ad Manager API
 * @return A list of line item errors
 */
function extractLineItemErrors(
  serverFault: AdManagerServerFault,
  lineItems: ad_manager.LineItem[],
): LineItemError[] {
  const fieldPathRegex = /^lineItem\[(\d+)\]\.([^;]*)?$/;

  const errors: LineItemError[] = [];

  serverFault.errors.forEach((error) => {
    const fieldPathMatch = error.fieldPath.match(fieldPathRegex);

    if (fieldPathMatch) {
      const [, lineItemIndex, fieldPath] = fieldPathMatch;

      const lineItem = lineItems[parseInt(lineItemIndex)];

      const lineItemError: LineItemError = {
        lineItemId: lineItem.id,
        errorString: `${error.errorString} @ lineItem[${lineItem.id}].${fieldPath};`,
      };

      errors.push(lineItemError);
    }
  });

  return errors;
}

/**
 * Returns true if the API error is retryable, false otherwise.
 * @param serverFault The error returned from the Ad Manager API
 * @return True if the error is retryable, false otherwise
 */
function isTransientAdManagerError(serverFault: AdManagerServerFault): boolean {
  const errorString = serverFault.errors[0].errorString;

  switch (errorString) {
    case 'CommonError.CONCURRENT_MODIFICATION':
    case 'QuotaError.EXCEEDED_QUOTA': {
      return true;
    }
    default:
      return false;
  }
}

/**
 * Given a list of line item IDs, retrieve the corresponding line items from
 * Ad Manager, apply a transformation to the line items, and then upload the
 * modified line items to Ad Manager. To ensure performance, these operations
 * are performed in batches.
 *
 * If any line items fail to upload due to configuration issues (e.g. inactive
 * key-value targeting, etc.), then those lines will be removed and the upload
 * will be resubmitted. If server failures occur (e.g. quota exceeded or
 * concurrency issues), then the upload will also be retried up to 8 times with
 * an increasing delay between attempts.
 * @param lineItemIds A complete list of all line item IDs to upload
 * @param offset The offset into the ID array to use for the request
 * @param limit The number of line items to handle
 * @param transformer The function to apply to the line items
 * @return A list of error messages for display
 */
function transformAndUploadLineItems(
  lineItemIds: number[],
  offset: number,
  limit: number,
  transformer: (lineItems: ad_manager.LineItem[]) => void,
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
): string[] {
  let lineItems: ad_manager.LineItem[] = [];
  let errorMessages: string[] = [];

  let attempt = 1;
  let lineItemsRetrieved = false;
  let uploadComplete = false;

  const lineItemIdsBatch = lineItemIds.slice(offset, offset + limit);

  do {
    try {
      if (!lineItemsRetrieved) {
        // Don't redownload line items if we already have them
        lineItems = getLineItemsWithIds(lineItemIdsBatch, adManagerHandler);

        lineItemsRetrieved = true;
      }

      transformer(lineItems);

      adManagerHandler.uploadLineItems(lineItems);

      uploadComplete = true;
    } catch (e) {
      if (e instanceof AdManagerServerFault) {
        if (isTransientAdManagerError(e)) {
          Logger.log('Error: Waiting ' + attempt + ' seconds for retry.');

          Utilities.sleep(attempt * 1000);
        } else {
          const lineItemErrors = extractLineItemErrors(e, lineItems);

          if (lineItemErrors.length > 0) {
            Logger.log('Error: Removing line items failures before retry.');
            // Remove any line items that failed to upload from the batch
            lineItems = lineItems.filter(
              (x) => !lineItemErrors.some((y) => y.lineItemId === x.id),
            );

            errorMessages.push(...lineItemErrors.map((x) => x.errorString));
          } else {
            throw e; // Error is not retryable and is not line item specific
          }
        }
      } else {
        throw e; // Rethrow unrecognized error
      }
    }

    attempt++;
  } while (!uploadComplete && attempt <= 8);

  updateUploadStatus(lineItemIds.length, limit);

  return errorMessages;
}

/**
 * Updates the progress bar with the number of line items uploaded. Due to the
 * concurrent execution of the upload process, we rely on the AppsScript lock
 * service to ensure that only one instance of this function is running at a
 * time.
 */
function updateUploadStatus(lineItemCount: number, limit: number) {
  const lock = LockService.getUserLock();

  try {
    lock.waitLock(5000);

    const taskProgress = getTaskProgress();

    setTaskProgress('Uploaded', taskProgress.current + limit, lineItemCount);
  } catch (e) {
    Logger.log('Could not obtain lock after 5 seconds.');
  } finally {
    lock.releaseLock();
  }
}

/**
 * Applies historical delivery pacing to all selected line items within the
 * active sheet.
 */
function applyHistorical(
  lineItemIds: number[],
  offset: number,
  limit: number,
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
): void {
  transformAndUploadLineItems(
    lineItemIds,
    offset,
    limit,
    (lineItems) => {
      adManagerHandler.applyHistoricalToLineItems(lineItems);
    },
    adManagerHandler,
  );
}

/**
 * Updates the line items with the custom delivery curve specified in the
 * active sheet.
 * @param lineItemIds A complete list of all line item IDs to upload
 * @param offset The offset into the ID array to use for the request
 * @param limit The number of line items to handle
 */
function uploadLineItems(
  lineItemIds: number[],
  offset: number,
  limit: number,
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
  sheetHandler: SheetHandler = getSheetHandler(),
): void {
  const curveTemplate = sheetHandler.getCurveTemplate();

  const errorMessages = transformAndUploadLineItems(
    lineItemIds,
    offset,
    limit,
    (lineItems) => {
      adManagerHandler.applyCurveToLineItems(lineItems, curveTemplate);
    },
    adManagerHandler,
  );

  if (errorMessages.length > 0) {
    throw new Error(errorMessages.join(','));
  }
}

global.onEdit = onEdit;
global.onOpen = onOpen;
global.callback = callback;
global.include = include;
global.showSidebar = showSidebar;

export const TEST_ONLY = {
  beginLoadLineItems,
  copyTemplate,
  getAdManagerHandler,
  getLineItemPreviews,
  getTaskProgress,
  loadLineItems,
  setTaskProgress,
};
