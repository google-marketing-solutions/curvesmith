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
 * @fileoverview Shared client-side code.
 */

/**
 * Describes the current and total progress values for the active task. Some
 * tasks may not opt to use this functionality, but long-running tasks can
 * communicate status back to a client script.
 */
export declare interface TaskProgress {
  current: number;
  total: number;
}

let startTime: Date, timerInterval: number;

/** Base class for UI elements. */
export class UIElementRenderer extends EventTarget {
  constructor() {
    super();
  }

  /**
   * Updates the UI to reflect the start of a task. Typically this is called
   * before a server script callback is initiated.
   */
  beginTask(message: string) {
    this.logStatus(message);

    this.startProgress();
  }

  /**
   * Updates the UI to reflect a task that has completed with an error.
   *
   * Typically this is called by a server script's `withFailureHandler`.
   */
  finishTaskWithFailure(error: Error) {
    this.stopProgress();

    const status = formatErrorMessage(error);

    this.logStatus(status, /* failure= */ true);
  }

  /**
   * Updates the UI to reflect a task that has completed successfully.
   *
   * Typically this is called by a server script's `withSuccessHandler`.
   */
  finishTaskWithSuccess(message: string) {
    this.stopProgress();

    this.logStatus(message);
  }

  /** Logs a status message to the UI. */
  logStatus(status: string, failure = false) {
    this.queryAndExecute<HTMLElement>('.status', (statusCell) => {
      if (failure) {
        statusCell.classList.replace('info', 'error');
      } else {
        statusCell.classList.replace('error', 'info');
      }

      statusCell.textContent = status;
    });
  }

  /**
   * Queries the DOM for all selector matches and executes a callback for each.
   */
  queryAndExecute<T extends Element>(
    selector: string,
    callback: (element: T) => void,
  ): void {
    const elements = document.querySelectorAll(selector);

    if (elements.length === 0) {
      console.warn(`Element not found or doesn't match selector: ${selector}`);
    } else {
      for (const element of elements) {
        if (element instanceof Element) {
          callback(element as T);
        }
      }
    }
  }

  /** Returns a formatted string of the time elapsed since a task started. */
  private formatTimeElapsed(): string {
    const difference = Date.now() - startTime.getTime();
    const totalSeconds = Math.round(difference / 1000);
    const mins = Math.floor(totalSeconds / 60);
    const secs = totalSeconds % 60;

    return `${String(mins).padStart(2, '0')}:${String(secs).padStart(2, '0')}`;
  }

  /** Initializes the progress bar and kicks off timer tracking. */
  private startProgress() {
    startTime = new Date();
    timerInterval = window.setInterval(() => this.updateProgress(), 1000);

    // Show a small progress bar to indicate that the script is running.
    this.queryAndExecute<HTMLElement>('.progress-bar', (progressBar) => {
      progressBar.style.width = '5%';
      progressBar.style.animation = 'flash 1s infinite alternate';
    });
  }

  /** Fills the progress bar and stops the timer. */
  private stopProgress() {
    window.clearInterval(timerInterval);

    google.script.run['callback']('clearTaskProgress');

    this.queryAndExecute<HTMLElement>('.progress-bar', (progressBar) => {
      progressBar.style.width = '100%';
      progressBar.style.animation = 'none';
    });

    this.queryAndExecute<HTMLElement>('.progress-text', (progressText) => {
      progressText.textContent = '';
    });
  }

  /**
   * Updates the progress bar and timer in the UI. Progress status is dependent
   * on a server callback (`getTaskProgress`) that provides feedback. Not all
   * server callbacks implement this feedback.
   */
  private updateProgress() {
    this.queryAndExecute<HTMLElement>('.timer', (timer) => {
      timer.textContent = 'Elapsed Time: ' + this.formatTimeElapsed();
    });

    google.script.run
      .withSuccessHandler((x) => this.updateProgressBar(x))
      ['callback']('getTaskProgress');
  }

  /** Updates the progress bar to reflect task status. */
  private updateProgressBar(taskProgress: TaskProgress) {
    if (taskProgress.total > 0) {
      this.queryAndExecute<HTMLElement>('.progress-bar', (progressBar) => {
        const relative = taskProgress.current / taskProgress.total;
        const progress = Math.min(relative * 100, 100);

        progressBar.style.width = `${progress}%`;
      });

      this.queryAndExecute<HTMLElement>('.progress-text', (progressText) => {
        progressText.textContent = `Progress: ${taskProgress.current} of ${taskProgress.total}`;
      });
    }
  }
}

/**
 * Returns a formatted error message based on the provided error. Most errors
 * are passed through as-is, but Ad Manager errors are formatted to include the
 * error code and, if available, the trigger.
 */
export function formatErrorMessage(error: Error): string {
  const adManagerMatch = error.message.match(
    /AdManagerServerFault: \[.*\.(.*) @ (?:\; trigger:'(.*?)')?\]/,
  );

  if (adManagerMatch) {
    const [, errorCode, trigger] = adManagerMatch;

    return trigger
      ? `Ad Manager Error: ${errorCode} (${trigger})`
      : `Ad Manager Error: ${errorCode}`;
  } else {
    return error.message;
  }
}
