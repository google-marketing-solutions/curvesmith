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
 * @fileoverview Handles interactions with a sheet in the active spreadsheet.
 */

import {CurveTemplate, GoalType, ScheduledEvent} from './custom_curve';

/** Represents a single row of line item metadata within a sheet. */
export interface LineItemRow {
  id: number;
  name: string;
  startDate: string;
  endDate: string;
  impressionGoal: number;
  selected: boolean;
}

/**
 * Handles reading and writing data relevant to custom curve creation within a
 * specific sheet.
 *
 * All date values are relative to the to the spreadsheet timezone.
 */
export class SheetHandler {
  /** Named range where the user can specify an ad unit ID. [Size: 1x1] */
  static readonly NAMED_RANGE_AD_UNIT_ID = 'AD_UNIT_ID';

  /** Named range where the user can specify the goal type. [Size: 1x1] */
  static readonly NAMED_RANGE_GOAL_TYPE = 'GOAL_TYPE';

  /** Named range where the user can specify line items. [Size: 5x?] */
  static readonly NAMED_RANGE_LINE_ITEMS = 'LINE_ITEMS';

  /** Named range where the user can specify custom events. [Size: 4x?]*/
  static readonly NAMED_RANGE_SCHEDULED_EVENTS = 'SCHEDULED_EVENTS';

  constructor(readonly sheet: GoogleAppsScript.Spreadsheet.Sheet) {}

  /**
   * Appends line item rows to the associated sheet within the designated named
   * range (`LINE_ITEMS`). Existing data will be preserved and, if necessary,
   * the named range will be expanded to accommodate the new data.
   * @param lineItems The line items to write to the sheet
   * @return The row index (1-based) where the new data was written
   */
  appendLineItems(lineItems: LineItemRow[]): number | undefined {
    if (lineItems.length > 0) {
      const appendRange = this.getAppendRange(
        SheetHandler.NAMED_RANGE_LINE_ITEMS,
        lineItems.length,
      );

      const lineItemValues = lineItems.map((lineItem) => [
        /* selected= */ false,
        /* id= */ lineItem.id,
        /* name= */ lineItem.name,
        /* startDate= */ lineItem.startDate,
        /* endDate= */ lineItem.endDate,
        /* impressionGoal= */ lineItem.impressionGoal,
      ]);

      const selectedColumnRange = this.sheet.getRange(
        /* row= */ appendRange.getRow(),
        /* column= */ appendRange.getColumn(),
        /* numRows= */ appendRange.getNumRows(),
        /* numColumns= */ 1,
      );

      // Add checkboxes to the "Selected" column
      const dataValidation = SpreadsheetApp.newDataValidation()
        .requireCheckbox()
        .setAllowInvalid(false)
        .setHelpText('The value of this cell must be true or false')
        .build();

      selectedColumnRange.setDataValidation(dataValidation);

      return appendRange.setValues(lineItemValues).getRow();
    }

    return undefined;
  }

  /** Clears all line item content from the associated sheet. */
  clearLineItems() {
    const lineItemsRange = this.getNamedRange(
      SheetHandler.NAMED_RANGE_LINE_ITEMS,
    );

    lineItemsRange.clearContent();
    lineItemsRange.removeCheckboxes();
  }

  /**
   * Returns the filtering ad unit ID from the associated sheet. This depends
   * on the existence of the `AD_UNIT_ID` named range.
   */
  getAdUnitId(): string {
    return String(this.getNamedValue(SheetHandler.NAMED_RANGE_AD_UNIT_ID));
  }

  /**
   * Returns a `CurveTemplate` instance representing the configuration data
   * present in the associated sheet.
   */
  getCurveTemplate(): CurveTemplate {
    const goalType = this.getGoalType();
    const events = this.getScheduledEvents();

    return new CurveTemplate(events, goalType);
  }

  /**
   * Returns the goal type from the associated sheet. This depends on the
   * existence of the `GOAL_TYPE` named range.
   */
  getGoalType(): GoalType {
    const goalType = this.getNamedValue(SheetHandler.NAMED_RANGE_GOAL_TYPE);

    return GoalType[goalType as keyof typeof GoalType];
  }

  /**
   * Returns all line item rows with the `selected` column set to `true` from
   * the associated sheet. This depends on the existence of the `LINE_ITEMS`
   * named range.
   */
  getSelectedLineItems(): LineItemRow[] {
    const lineItemsRange = this.getNamedRange(
      SheetHandler.NAMED_RANGE_LINE_ITEMS,
    );

    const lineItemRows: LineItemRow[] = [];

    for (const row of lineItemsRange.getValues()) {
      const [selected, idText, name, startDate, endDate, goalText] = row;

      if (!selected) {
        continue;
      }

      const id = parseFloat(idText);
      const impressionGoal = parseFloat(goalText);

      if (isNaN(id) || isNaN(impressionGoal)) {
        continue;
      }

      lineItemRows.push({
        id,
        name,
        startDate,
        endDate,
        impressionGoal,
        selected,
      });
    }

    return lineItemRows;
  }

  /**
   * Returns an array of scheduled events from the associated sheet. This
   * depends on the existence of the `SCHEDULED_EVENTS` named range.
   */
  getScheduledEvents(): ScheduledEvent[] {
    // [[Start, End, Goal Percent, Title],..]
    const rangeValues = this.getNamedRange(
      SheetHandler.NAMED_RANGE_SCHEDULED_EVENTS,
    ).getValues();

    const events: ScheduledEvent[] = [];

    for (const eventRow of rangeValues) {
      const [start, end, goalPercent, title] = eventRow;

      if (start.length === 0) {
        continue; // Ignore empty rows
      }

      // Both dates will be relative to the spreadsheet timezone.
      if (!(start instanceof Date || end instanceof Date)) {
        throw new Error('Scheduled event start and end must both be dates');
      }

      events.push(
        new ScheduledEvent(start, end, Number(goalPercent), String(title)),
      );
    }

    return events;
  }

  /**
   * Given a named range, identifies the first empty row and returns a sub-range
   * where the provided number of rows (e.g. `count`) should be appended. If the
   * named range is already fully populated, then it will be expanded by the
   * provided number of rows.
   * @param name The name of the range that will receive new rows
   * @param count The number of rows that will be appended
   */
  private getAppendRange(
    name: string,
    count: number,
  ): GoogleAppsScript.Spreadsheet.Range {
    const namedRange = this.getNamedRange(name);

    const values = namedRange.getValues();

    // Find the index of the first empty row
    const emptyRowIndex = values.findIndex((r) => r.every((c) => c === ''));

    if (emptyRowIndex < 0) {
      return this.sheet.getRange(
        /* row= */ namedRange.getLastRow() + 1,
        /* column= */ namedRange.getColumn(),
        /* numRows= */ count,
        /* numColumns= */ namedRange.getNumColumns(),
      );
    } else {
      return this.sheet.getRange(
        /* row= */ namedRange.getRow() + emptyRowIndex,
        /* column= */ namedRange.getColumn(),
        /* numRows= */ count,
        /* numColumns= */ namedRange.getNumColumns(),
      );
    }
  }

  /**
   * Returns the sheet range associated with the provided name if it exists
   * locally within the associated sheet.
   *
   * Each named range within a template sheet can be referenced locally by
   * prepending the sheet name to the range name (e.g. 'Sheet1!RANGE_NAME').
   * This function explicitly only looks for a local named range and will throw
   * an error if a match could not be found.
   * @param name The name of the range to read
   * @throws An error if the named range does not exist within this sheet
   */
  private getNamedRange(name: string): GoogleAppsScript.Spreadsheet.Range {
    const sheetName = this.sheet.getName();
    const localName = `'${sheetName}'!${name}`;

    const spreadsheet = this.sheet.getParent();
    const localRange = spreadsheet.getRangeByName(localName);

    if (!localRange || localRange.getSheet().getName() !== sheetName) {
      throw new RangeError(`${name} range does not exist`);
    }
    return localRange;
  }

  /**
   * Returns the value of the first cell associated with the provided named
   * range if it exists.
   * @param name The name of the range to read
   * @throws An error if the named range does not exist
   */
  private getNamedValue(name: string): any {
    return this.getNamedRange(name).getValue();
  }
}

/** Handles top-level interactions with the active spreadsheet. */
export class SpreadsheetHandler {
  /** Named range where the user can enter the Ad Manager API version. */
  static readonly NAMED_RANGE_AD_MANAGER_API_VERSION = 'API_VERSION';

  /** Named range where the user can enter the Ad Manager network ID. */
  static readonly NAMED_RANGE_AD_MANAGER_NETWORK_ID = 'NETWORK_ID';

  /** Named range where the user can opt to include an approve all button. */
  static readonly NAMED_RANGE_SHOW_APPROVE_ALL = 'SHOW_APPROVE_ALL';

  /** Named range where the user can enter the name of a template sheet. */
  static readonly NAMED_RANGE_TEMPLATE_SHEET_NAME = 'TEMPLATE_SHEET_NAME';

  constructor(readonly spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet) {}

  /** Creates a copy of the template sheet and activates the copy. */
  copyTemplate() {
    const templateName = String(
      this.getNamedValue(SpreadsheetHandler.NAMED_RANGE_TEMPLATE_SHEET_NAME),
    );
    const templateSheet = this.spreadsheet.getSheetByName(templateName);

    if (templateSheet) {
      templateSheet.copyTo(this.spreadsheet).activate()
    } else {
      throw new Error('Template sheet does not exist');
    }
  }

  /** Returns the Ad Manager API version from the associated spreadsheet. */
  getApiVersion(): string {
    return String(
      this.getNamedValue(SpreadsheetHandler.NAMED_RANGE_AD_MANAGER_API_VERSION),
    );
  }

  /** Returns the Ad Manager network ID from the associated spreadsheet. */
  getNetworkId(): string {
    return String(
      this.getNamedValue(SpreadsheetHandler.NAMED_RANGE_AD_MANAGER_NETWORK_ID),
    );
  }

  /** Returns the 'Show Approve All' setting from the associated spreadsheet. */
  getShowApproveAll(): boolean {
    return Boolean(
      this.getNamedValue(SpreadsheetHandler.NAMED_RANGE_SHOW_APPROVE_ALL),
    );
  }

  /**
   * Sets the spreadsheet timezone to the provided value.
   *
   * This function can be used to ensure that the spreadsheet timezone is
   * consistent with a specific Ad Manager network configuration. Absent this
   * explicit assurance, time handling may be inconsistent with user intent.
   * @param timeZoneId An IANA timezone ID value like "America/New_York"
   */
  updateSpreadsheetTimeZone(timeZoneId: string) {
    this.spreadsheet.setSpreadsheetTimeZone(timeZoneId);
  }

  /**
   * Returns the value of the first cell associated with the provided named
   * range if it exists.
   *
   * This function looks for an exact match of the provided name and will throw
   * an error if a match could not be found. Typical usage will be to lookup
   * configuration values like the Ad Manager API version or network ID.
   *
   * By comparison, the {@link SheetHandler} class provides a similar method
   * that will explicitly only look for a named range defined locally within the
   * associated sheet.
   * @param name The name of the range to read
   * @throws An error if the named range does not exist
   */
  private getNamedValue(name: string): unknown {
    const namedRange = this.spreadsheet.getRangeByName(name);

    if (!namedRange) {
      throw new RangeError(`${name} range does not exist`);
    } else {
      return namedRange.getValue();
    }
  }
}
