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
 * @fileoverview Client-side code for the preview and upload dialogs.
 */

import "./common.css";
import "./upload.css";

import {sanitizeHtml} from 'safevalues';
import {safeElement} from 'safevalues/dom';
import {UIElementRenderer} from './common';

/**
 * Describes a single goal segment of a custom curve. This corresponds to the
 * curve data that would be uploaded into Ad Manager. Notably `startDate` is a
 * string because Date objects are not supported by Apps Script callbacks.
 */
export declare interface CurveGoal {
  description: string;
  startDate: string;
  goalPercent: number;
  impressionGoal: number;
}

/**
 * Encapsulates a line item and a preview of the custom curve that would be
 * uploaded into Ad Manager. Notably `startDate` and `endDate` strings because
 * Date objects are not supported by Apps Script callbacks.
 */
export declare interface LineItemPreview {
  id: number;
  name: string;
  startDate: string;
  endDate: string;
  impressionGoal: number;
  curveGoals: CurveGoal[];
}

/** Encapsulates server-side settings to avoid unnecessary callbacks. */
export declare interface DialogSettings {
  /** The Ad Manager network ID associated with the preview. */
  networkId: string;

  /** Indicates whether to display only the preview. */
  previewOnly: boolean;

  /** Indicates whether to display the batch approval button. */
  showApproveAll: boolean;
}

/** A set of values that is always sorted. */
class SortedSet<T> {
  private set: Set<T> = new Set();

  constructor(readonly compareFn: (a: T, b: T) => number) {}

  /** Adds a unique value to the sorted set. */
  add(value: T) {
    this.set = new Set<T>(Array.from(this.set.add(value)).sort(this.compareFn));
  }

  /** Removes a value from the sorted set. */
  delete(value: T) {
    this.set.delete(value);
  }

  /** Returns true if the sorted set contains the value. */
  has(value: T) {
    return this.set.has(value);
  }

  /** Returns the number of values in the sorted set. */
  get size() {
    return this.set.size;
  }

  /** Returns an ordered collection of values from the sorted set. */
  get values(): T[] {
    return Array.from(this.set);
  }
}

let renderer: UploadRenderer;

/**
 * This class manages the rendering of and interaction with UI elements of the
 * preview and upload dialogs. The preview dialog only presents the user with a
 * single stage to view the hypothetical application of a configured custom
 * curve to selected line items. The upload dialog incorporates two additional
 * stages in order to support upload into Ad Manager and then display failures
 * if any occur.
 * @extends UIElementRenderer
 */
export class UploadRenderer extends UIElementRenderer {
  /** A map of stage names to their corresponding DOM element selectors. */
  private static readonly STAGE_CLASSES: Record<string, string> = {
    'preview': '.stage1-preview',
    'upload': '.stage2-upload',
    'failure': '.stage3-failure',
  };

  /** Number of line item IDs to display per table row. */
  private static readonly LINE_ITEM_IDS_PER_ROW = 5;

  /** Height of the reserved title space at the top of the dialog. */
  private static readonly DIALOG_HEADER_HEIGHT = 32;

  /** Maximum height of the dialog. */
  private static readonly DIALOG_MAXIMUM_HEIGHT = 500;

  /** A set of line item IDs that have been approved by the user for upload. */
  readonly approvedLineItemIds = new SortedSet<number>((a, b) => a - b);

  currentItemIndex = 0;
  currentStage = 'preview';

  constructor(
    readonly lineItems: LineItemPreview[],
    readonly settings: DialogSettings,
  ) {
    super();

    this.initializeNavigationButtons();

    if (!settings.previewOnly) {
      this.initializeUploadButtons();
    }

    this.initializeClickHandlers();
    this.initializeKeyboardShortcuts();
  }

  override beginTask(message: string) {
    super.beginTask(message);

    // Hide the buttons while the upload is running
    this.queryAndExecute<HTMLElement>('.buttons', (buttons) => {
      buttons.style.display = 'none';
    });

    this.showCloseButton(/* enabled= */ false);
  }

  override finishTaskWithFailure(error: Error) {
    super.finishTaskWithFailure(error);

    this.parseLineItemFailures(error.message);

    this.showCloseButton(/* enabled= */ true);
  }

  override finishTaskWithSuccess(message: string) {
    super.finishTaskWithSuccess(message);

    this.showCloseButton(/* enabled= */ true);
  }

  /** Approves all line items and skips to the upload stage. */
  approveAll() {
    for (const lineItem of this.lineItems) {
      this.approvedLineItemIds.add(lineItem.id);
    }

    this.syncButtonsWithState();
    this.showUploadStage();
  }

  displayFirst() {
    this.displayLineItem(0);
  }

  displayLast() {
    this.displayLineItem(this.lineItems.length - 1);
  }

  displayNext() {
    if (this.currentItemIndex < this.lineItems.length - 1) {
      this.displayLineItem(this.currentItemIndex + 1);
    }
  }

  displayPrevious() {
    if (this.currentItemIndex > 0) {
      this.displayLineItem(this.currentItemIndex - 1);
    }
  }

  /** Returns the currently displayed line item. */
  getCurrentLineItem(): LineItemPreview {
    return this.lineItems[this.currentItemIndex];
  }

  /**
   * Displays the first stage of the upload dialog. This stage consists of the
   * user previewing each custom curve and marking each line item approved prior
   * to confirming the upload.
   */
  showPreviewStage() {
    this.showStage('preview');
  }

  /**
   * Displays the second stage of the upload dialog, which presents the user
   * with a list of line items that they have approved and a confirmation button
   * to actually submit the upload to Ad Manager.
   */
  showUploadStage() {
    this.queryAndExecute<HTMLTableElement>('.line-items', (table) => {
      this.clearTable(table);

      let cellCount = 0;
      let currentRow = table.insertRow();

      this.approvedLineItemIds.values.forEach((lineItemId) => {
        this.renderSafeHtml(
          /* element= */ currentRow.insertCell(),
          /* html= */ this.createLineItemLink(lineItemId),
        );

        cellCount++;

        if (cellCount === UploadRenderer.LINE_ITEM_IDS_PER_ROW) {
          currentRow = table.insertRow();
          cellCount = 0;
        }
      });
    });

    this.showStage('upload');
  }

  /**
   * Displays the third stage of the upload dialog. This stage displays a list
   * of errors returned by Ad Manager if the upload failed.
   */
  showFailureStage() {
    this.showStage('failure');
  }

  /** Toggles the current line item between approved and unapproved. */
  toggleCurrent() {
    const currentItemId = this.getCurrentLineItem().id;

    if (this.approvedLineItemIds.has(currentItemId)) {
      this.approvedLineItemIds.delete(currentItemId);
    } else {
      this.approvedLineItemIds.add(currentItemId);
    }

    this.syncButtonsWithState();
  }

  /** Clears all rows from the given table. */
  protected clearTable(table: HTMLTableElement) {
    while (table.rows.length > 0) {
      table.deleteRow(0); // Remove the first row until the body is empty
    }
  }

  /** Creates a button with the given class name and text label. */
  protected createButton(className: string, label: string): HTMLButtonElement {
    const button = document.createElement('button');
    button.className = className;
    button.textContent = label;
    return button;
  }

  /** Returns a hyperlink to the Ad Manager line item. */
  protected createLineItemLink(lineItemId: number) {
    const adManagerUrl = `https://admanager.google.com/${this.settings.networkId}#delivery/line_item/detail/line_item_id=${lineItemId}`;

    return `<a href="${adManagerUrl}" target="_blank" rel="noopener noreferrer">${lineItemId}</a>`;
  }

  /** Safely renders the provided HTML into the given DOM element. */
  protected renderSafeHtml(element: HTMLElement, html: string) {
    const safeHtml = sanitizeHtml(html);

    safeElement.setInnerHtml(element, safeHtml);
  }

  /** Adds curve preview data to a table within the dialog. */
  private addCurvePreview(curveGoals: CurveGoal[]) {
    this.queryAndExecute<HTMLTableElement>('.curve', (table) => {
      this.clearTable(table);

      const headerRow = table.insertRow();
      const headers = ['Description', 'Start', 'Goal %', 'Impressions'];

      headers.forEach((header) => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
      });

      curveGoals.forEach((curveGoal) => {
        const row = table.insertRow();
        const descriptionCell = row.insertCell();
        descriptionCell.classList.add('description');
        descriptionCell.textContent = curveGoal.description;

        const startCell = row.insertCell();
        startCell.textContent = this.formatDate(curveGoal.startDate);

        const goalPercentCell = row.insertCell();
        goalPercentCell.textContent = curveGoal.goalPercent.toFixed(3) + '%';

        const impressionGoalCell = row.insertCell();
        const roundedImpressionGoal = Math.round(curveGoal.impressionGoal);
        impressionGoalCell.textContent = roundedImpressionGoal.toLocaleString();
      });
    });
  }

  /**
   * Adjusts the dialog height to better accommodate the current stage by
   * eliminating unnecessary empty space. There is a maximum height that the
   * dialog can reach, but it can be reduced to a smaller size if the stage
   * doesn't require the full height.
   */
  private adjustDialogHeight() {
    const stageClass = UploadRenderer.STAGE_CLASSES[this.currentStage];

    let height = UploadRenderer.DIALOG_HEADER_HEIGHT;

    // Sums the height from each display block of the stage.
    this.queryAndExecute<HTMLElement>(stageClass, (stage) => {
      height += stage.getBoundingClientRect().height;
    });

    const dialogHeight = Math.min(height, UploadRenderer.DIALOG_MAXIMUM_HEIGHT);

    google.script.host.setHeight(dialogHeight);
  }

  /** Displays a line item preview within the dialog. */
  private displayLineItem(index: number = 0) {
    this.currentItemIndex = index;

    const lineItem = this.getCurrentLineItem();

    // Reset the scroll position of the main container.
    this.queryAndExecute<HTMLElement>('main', (main) => {
      main.scrollTop = 0;
    });

    this.queryAndExecute<HTMLButtonElement>('.current', (button) => {
      button.textContent = `${index + 1} of ${this.lineItems.length}`;
    });

    this.queryAndExecute<HTMLTableElement>('.line-item', (table) => {
      this.clearTable(table);

      const addTableRow = (label: string, value: string) => {
        const row = table.insertRow();
        const labelCell = row.insertCell();
        const valueCell = row.insertCell();

        labelCell.textContent = `${label}:`;
        labelCell.classList.add('label');

        this.renderSafeHtml(valueCell, value);

        return row;
      };

      addTableRow('Line Item', this.createLineItemLink(lineItem.id));
      addTableRow('Name', lineItem.name);
      addTableRow('Start', this.formatDate(lineItem.startDate));
      addTableRow('End', this.formatDate(lineItem.endDate));
      addTableRow('Goal', lineItem.impressionGoal.toLocaleString());

      this.addCurvePreview(lineItem.curveGoals);
    });

    this.syncButtonsWithState();
  }

  /**
   * Formats a date string into a human-readable date. Typically this function
   * will be used to reformat ISO 8601 date strings.
   */
  private formatDate(value: string): string {
    return new Date(value).toLocaleString();
  }

  /** Initializes a single click handler for all buttons in the dialog. */
  private initializeClickHandlers() {
    document.addEventListener('click', (event) => {
      const target = event.target as HTMLElement;

      switch (target.closest('button')?.className) {
        case 'approve-all':
          return this.approveAll();
        case 'back':
          return this.showPreviewStage();
        case 'cancel':
        case 'close':
          return closeDialog();
        case 'confirm':
          return confirmUpload();
        case 'first':
          return this.displayFirst();
        case 'last':
          return this.displayLast();
        case 'next':
          return this.displayNext();
        case 'prev':
          return this.displayPrevious();
        case 'toggle approved':
        case 'toggle unapproved':
          return this.toggleCurrent();
        case 'upload':
          return this.showUploadStage();
      }
    });
  }

  /** Initializes keyboard handling for the navigation actions. */
  private initializeKeyboardShortcuts() {
    document.addEventListener('keydown', (event) => {
      switch (event.key) {
        case 'ArrowLeft':
          this.displayPrevious();
          break;
        case 'ArrowRight':
          this.displayNext();
          break;
        case 'Enter':
          this.toggleCurrent();
          break;
      }
    });
  }

  /** Creates and adds the navigation buttons to the dialog. */
  private initializeNavigationButtons() {
    this.queryAndExecute<HTMLElement>('.nav', (element) => {
      element.append(
        this.createButton(/* className= */ 'first', /* label= */ '<<'),
        this.createButton(/* className= */ 'prev', /* label= */ '<'),
        this.createButton(/* className= */ 'current', /* label= */ 'Current'),
        this.createButton(/* className= */ 'next', /* label= */ '>'),
        this.createButton(/* className= */ 'last', /* label= */ '>>'),
      );
    });
  }

  /** Creates and adds the upload buttons to the dialog. */
  private initializeUploadButtons() {
    this.queryAndExecute<HTMLElement>('.act', (element) => {
      if (this.settings.showApproveAll) {
        element.append(
          this.createButton(
            /* className= */ 'approve-all',
            /* label= */ 'Approve All',
          ),
        );
      }

      element.append(
        this.createButton(
          /* className= */ 'toggle unapproved',
          /* label= */ 'Unapproved',
        ),
        this.createButton(/* className= */ 'upload', /* label= */ 'Upload'),
      );
    });
  }

  /**
   * Unfortunately Apps Script doesn't return the actual error object in the
   * failure handler; instead, the error message is packaged in a ScriptError
   * object. This method parses the error message to extract the actual error
   * details specific to each line item.
   *
   * If the error message doesn't match the expected format, then a generic
   * server fault was likely encountered and we'll just display that error
   * without navigating to the result stage.
   */
  private parseLineItemFailures(errorMessage: string) {
    const lineItemIds = this.approvedLineItemIds.values;

    // Matches line item details from AdManagerServerFault errors
    const lineItemRegex = /([^\[\.,\s]*) @ lineItem\[(\d+)\]\.([^;]*)?/g;

    const matches = errorMessage.matchAll(lineItemRegex);

    if (matches) {
      let errorCount = 0;
      for (const match of matches) {
        const [, errorType, lineItemIndex, fieldPath] = match;

        const lineItemId = lineItemIds[parseInt(lineItemIndex)];

        this.queryAndExecute<HTMLTableElement>('.errors', (table) => {
          const row = table.insertRow();
          row.insertCell().textContent = String(lineItemId);
          row.insertCell().textContent = errorType;
          row.insertCell().textContent = fieldPath;
        });
        errorCount++;
      }

      if (errorCount > 0) {
        this.showFailureStage();
      }
    }
  }

  private showCloseButton(enabled: boolean) {
    this.queryAndExecute<HTMLElement>('div.close', (div) => {
      div.style.display = 'block';
    });

    this.queryAndExecute<HTMLButtonElement>('button.close', (button) => {
      button.disabled = !enabled;
    });
  }

  /**
   * Used to show and hide stage elements in the DOM as a user navigates through
   * the upload dialog. Only one stage should be visible at a time.
   */
  private showStage(stageName: string) {
    for (const element in UploadRenderer.STAGE_CLASSES) {
      const stageClass = UploadRenderer.STAGE_CLASSES[element];

      this.queryAndExecute<HTMLElement>(stageClass, (e) => {
        e.style.display = element === stageName ? 'block' : 'none';
      });
    }

    this.currentStage = stageName;
    this.adjustDialogHeight();
  }

  /** Updates buttons to reflect state changes. */
  private syncButtonsWithState() {
    const currentItemId = this.getCurrentLineItem().id;

    // Update the toggle button to reflect the current approval status.
    this.queryAndExecute<HTMLButtonElement>('.toggle', (button) => {
      if (this.approvedLineItemIds.has(currentItemId)) {
        button.classList.replace('unapproved', 'approved');
        button.textContent = 'Approved';
      } else {
        button.classList.replace('approved', 'unapproved');
        button.textContent = 'Unapproved';
      }
    });

    // Update the upload button to reflect the number of approved line items.
    this.queryAndExecute<HTMLButtonElement>('.upload', (button) => {
      if (this.approvedLineItemIds.size > 0) {
        button.textContent = `Upload (${this.approvedLineItemIds.size})`;
        button.disabled = false;
      } else {
        button.textContent = 'Upload';
        button.disabled = true;
      }
    });

    // Disable the first button if the user is on the first line item.
    this.queryAndExecute<HTMLButtonElement>('.first', (firstButton) => {
      firstButton.disabled = this.currentItemIndex === 0;
    });

    // Disable the last button if the user is on the last line item.
    this.queryAndExecute<HTMLButtonElement>('.last', (lastButton) => {
      lastButton.disabled = this.currentItemIndex === this.lineItems.length - 1;
    });

    // Disable the next button if the user is on the last line item.
    this.queryAndExecute<HTMLButtonElement>('.next', (nextButton) => {
      nextButton.disabled = this.currentItemIndex >= this.lineItems.length - 1;
    });

    // Disable the previous button if the user is on the first line item.
    this.queryAndExecute<HTMLButtonElement>('.prev', (previousButton) => {
      previousButton.disabled = this.currentItemIndex <= 0;
    });
  }
}

function closeDialog() {
  google.script.host.close();
}

/** Sends the list of approved line item IDs back to the server script. */
function confirmUpload() {
  renderer.beginTask('Starting upload, please wait...');

  google.script.run
    .withSuccessHandler(() =>
      renderer.finishTaskWithSuccess('Upload complete.'),
    )
    .withFailureHandler((e) => renderer.finishTaskWithFailure(e))
    ['callback']('uploadLineItems', [renderer.approvedLineItemIds.values]);
}

/**
 * This helper function receives the result collection from the server and
 * displays the first line item. We are passing parameters into this client-side
 * code through scriplets in order to avoid a server roundtrip.
 */
function initialize(lineItems: LineItemPreview[], settings: DialogSettings) {
  renderer = new UploadRenderer(lineItems, settings);

  renderer.displayFirst();
  renderer.showPreviewStage();
}

global.initialize = initialize;
