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
 * @fileoverview Client-side code for the sidebar.
 */

import "./common.css";
import "./sidebar.css";

import {UIElementRenderer} from './common';

/**
 * Manages the rendering and dynamic updates of UI elements within an Apps
 * Script sidebar. This class is responsible for handling user interactions
 * within the sidebar environment.
 *
 * It interacts with the server-side Apps Script code to trigger actions and
 * update the sidebar's contents based on user input or other events.
 * @extends UIElementRenderer
 */
class SidebarRenderer extends UIElementRenderer {
  constructor() {
    super();
  }

  /**
   * Updates the sidebar UI to reflect the beginning of a task by disabling the
   * toolbar to prevent user interaction.
   */
  override beginTask(message: string) {
    this.disableToolbar();

    super.beginTask(message);
  }

  /**
   * Updates the sidebar UI to reflect the failed completion of a task. The
   * toolbar is re-enabled to allow user interaction.
   */
  override finishTaskWithFailure(error: Error) {
    super.finishTaskWithFailure(error);

    this.enableToolbar();
  }

  /**
   * Updates the sidebar UI to reflect the successful completion of a task. The
   * toolbar is re-enabled to allow user interaction.
   */
  override finishTaskWithSuccess(message = 'Task') {
    super.finishTaskWithSuccess(message);

    this.enableToolbar();
  }

  /** Prevents user from clicking on the toolbar while the script is running. */
  private disableToolbar() {
    this.queryAndExecute<HTMLButtonElement>('button', (button) => {
      button.disabled = true;
    });
  }

  /** Re-enables the toolbar after the script has finished running. */
  private enableToolbar() {
    this.queryAndExecute<HTMLButtonElement>('button', (button) => {
      button.disabled = false;
    });
  }
}

function initializeSpreadsheet() {
  renderer.beginTask('Initializing spreadsheet, please wait...');

  google.script.run
    .withSuccessHandler(() =>
      renderer.finishTaskWithSuccess('Spreadsheet initialized.'),
    )
    .withFailureHandler((e) => renderer.finishTaskWithFailure(e))
    ['callback']('initializeSpreadsheet');
}

export function applyHistorical() {
  renderer.beginTask('Applying historical, please wait...');

  google.script.run
    .withSuccessHandler(() =>
      renderer.finishTaskWithSuccess(
        'Delivery pacing source reset to historical.',
      ),
    )
    .withFailureHandler((e) => renderer.finishTaskWithFailure(e))
    ['callback']('applyHistorical');
}

export function copyTemplate() {
  renderer.beginTask('Creating new template, please wait...');

  google.script.run
    .withSuccessHandler(() =>
      renderer.finishTaskWithSuccess('New template created.'),
    )
    .withFailureHandler((e) => renderer.finishTaskWithFailure(e))
    ['callback']('copyTemplate');
}

export function loadLineItems() {
  renderer.beginTask('Retrieving line items, please wait...');

  const PAGE_SIZE = 50; // Number of line items to request per page

  let oneOrMoreFailures = false;
  let taskCompletionCount = 0;

  google.script.run
    .withSuccessHandler((totalLineItemCount: number) => {
      const requests = Math.ceil(totalLineItemCount / PAGE_SIZE);

      for (let i = 0; i < requests; i++) {
        google.script.run
          .withSuccessHandler(() => {
            if (!oneOrMoreFailures && ++taskCompletionCount === requests) {
              renderer.finishTaskWithSuccess('Line items retrieved.');
            }
          })
          .withFailureHandler((e) => {
            oneOrMoreFailures = true;
            renderer.finishTaskWithFailure(e);
          })
          ['callback']('loadLineItems', [i * PAGE_SIZE, PAGE_SIZE]);
      }
    })
    .withFailureHandler((e) => renderer.finishTaskWithFailure(e))
    ['callback']('beginLoadLineItems');
}

export function showPreviewDialog() {
  renderer.beginTask('Generating curve preview(s), please wait...');

  google.script.run
    .withSuccessHandler(() =>
      renderer.finishTaskWithSuccess('Preview(s) generated.'),
    )
    .withFailureHandler((e) => renderer.finishTaskWithFailure(e))
    ['callback']('showPreviewDialog');
}

export function showUploadDialog() {
  renderer.beginTask('Generating curve previews(s), please wait...');

  google.script.run
    .withSuccessHandler(() =>
      renderer.finishTaskWithSuccess('Preview(s) generated.'),
    )
    .withFailureHandler((e) => renderer.finishTaskWithFailure(e))
    ['callback']('showUploadDialog');
}

const renderer = new SidebarRenderer();

initializeSpreadsheet();

global.applyHistorical = applyHistorical;
global.copyTemplate = copyTemplate;
global.loadLineItems = loadLineItems;
global.showPreviewDialog = showPreviewDialog;
global.showUploadDialog = showUploadDialog;
