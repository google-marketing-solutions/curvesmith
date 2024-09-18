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
 * The number of objects to request at a time. This value was empirically
 * determined to be the optimal tradeoff between UX and performance.
 */
const AD_MANAGER_API_SUGGESTED_PAGE_SIZE = 50;

/**
 * Applies historical delivery pacing to all selected line items within the
 * active sheet.
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

  try {
    transformAndUploadLineItems(
      lineItemIds,
      (lineItems) => {
        adManagerHandler.applyHistoricalToLineItems(lineItems);
      },
      adManagerHandler,
    );
  } catch (error) {
    throw new Error('Partial completion');
  }
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
  let results: am_handler.LineItemDtoPage;

  do {
    results = adManagerHandler.getLineItemDtoPage(
      filter,
      offset,
      AD_MANAGER_API_SUGGESTED_PAGE_SIZE,
    );

    const lineItemRows: LineItemRow[] = [];

    for (const lineItem of results.values) {
      lineItemRows.push({
        selected: false,
        id: lineItem.id,
        name: lineItem.name,
        startDate: lineItem.startDate,
        endDate: lineItem.endDate,
        impressionGoal: lineItem.impressionGoal,
      });
    }

    sheetHandler.appendLineItems(lineItemRows);

    offset += AD_MANAGER_API_SUGGESTED_PAGE_SIZE;

    setTaskProgress('Retrieved', offset, offset);
  } while (!results.endOfResults);
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

/**
 * Extracts metadata about line items that failed to upload.
 * @param serverFault The error returned from the Ad Manager API
 * @param lineItems The line items that were sent to the Ad Manager API
 * @param failureIdList A modified set of line item IDs that failed to upload
 * @param errorMessages A modified list of error messages for display
 */
function identifyLineItemFailures(
  serverFault: AdManagerServerFault,
  lineItems: ad_manager.LineItem[],
  failureIdList: Set<number>,
  errorMessages: string[],
) {
  const fieldPathRegex = /^lineItem\[(\d+)\]\.([^;]*)?$/;

  serverFault.errors.forEach((error) => {
    const fieldPathMatch = error.fieldPath.match(fieldPathRegex);

    if (fieldPathMatch) {
      const [, lineItemIndex, fieldPath] = fieldPathMatch;

      const lineItem = lineItems[parseInt(lineItemIndex)];

      // Cache line items that are failing
      failureIdList.add(lineItem.id);

      // Cache error messages for display upon completion
      errorMessages.push(
        `${error.errorString} @ lineItem[${lineItem.id}].${fieldPath};`,
      );
    }
  });
}

/**
 * Given a list of line item IDs, retrieve the corresponding line items from
 * Ad Manager, apply a transformation to the line items, and then upload the
 * modified line items to Ad Manager. To ensure performance, these operations
 * are performed in batches. If any errors occur during a batch, the user is
 * given the option to exclude the failing line items and try again. This option
 * will only be presented once, after which the remaining line items will be
 * uploaded without further prompting.
 */
function transformAndUploadLineItems(
  lineItemIds: number[],
  transformer: (lineItems: ad_manager.LineItem[]) => void,
  adManagerHandler: am_handler.AdManagerHandler = getAdManagerHandler(),
): void {
  const batchSize = AD_MANAGER_API_SUGGESTED_PAGE_SIZE;

  const RETRY_MESSAGE =
    'One or more errors occurred during an upload batch. ' +
    'Exclude the problematic line items and try again?';

  let offset = 0;
  let firstFailure = true;
  let errorMessages: string[] = [];

  do {
    const batchOffset = offset + batchSize;
    const lineItemIdsBatch = lineItemIds.slice(offset, batchOffset);

    setTaskProgress('Retrieving', offset, lineItemIds.length);

    const lineItems = getLineItemsWithIds(lineItemIdsBatch, adManagerHandler);

    setTaskProgress('Retrieved', batchOffset, lineItemIds.length);

    transformer(lineItems);

    setTaskProgress('Uploading', batchOffset, lineItemIds.length);

    const failureIdList = new Set<number>();

    try {
      adManagerHandler.uploadLineItems(lineItems);
    } catch (e) {
      if (e instanceof AdManagerServerFault) {
        identifyLineItemFailures(e, lineItems, failureIdList, errorMessages);

        // If no lines were identified in the error, then the error is
        // unexpected and should be rethrown.
        if (failureIdList.size === 0) {
          throw e;
        }

        if (firstFailure && shouldCancelAction(RETRY_MESSAGE)) {
          break;
        } else {
          firstFailure = false;

          // Remove any line items that failed to upload from the batch
          const filteredLineItems = lineItems.filter(
            (lineItem) => !failureIdList.has(lineItem.id),
          );

          // Attempt to upload the remaining line items
          adManagerHandler.uploadLineItems(filteredLineItems);
        }
      }
    } finally {
      offset += lineItemIdsBatch.length;

      setTaskProgress('Uploaded', offset, lineItemIds.length);
    }
  } while (offset < lineItemIds.length);

  if (errorMessages.length > 0) {
    throw new Error(errorMessages.join(','));
  }
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

  transformAndUploadLineItems(
    lineItemIds,
    (lineItems) => {
      adManagerHandler.applyCurveToLineItems(lineItems, curveTemplate);
    },
    adManagerHandler,
  );
}

global.onEdit = onEdit;
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
