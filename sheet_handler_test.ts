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

import {GoalType, ScheduledEvent} from './custom_curve';
import {LineItemRow, SheetHandler} from './sheet_handler';

const MOCK_LINE_ITEM_ROW: LineItemRow = {
  id: 12345678,
  name: 'Line Item Name',
  startDate: '2024-01-01 00:00:00',
  endDate: '2024-12-31 00:00:00',
  impressionGoal: 1000,
  selected: false,
};

describe('SheetHandler', () => {
  let sheetHandler: SheetHandler;
  let sheetMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Sheet>;
  let spreadsheetAppMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.SpreadsheetApp>;
  let spreadsheetMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Spreadsheet>;
  let namedRangeMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.NamedRange>;
  let rangeMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Range>;

  beforeAll(() => {
    // SpreadsheetApp is unavailable outside of the Apps Script environment.
    spreadsheetAppMock = jasmine.createSpyObj('SpreadsheetApp', [
      'newDataValidation',
    ]);
    const dataValidationMock =
      jasmine.createSpyObj<GoogleAppsScript.Spreadsheet.DataValidation>(
        'DataValidation',
        ['getHelpText'],
      );
    const dataValidatorBuilderMock =
      jasmine.createSpyObj<GoogleAppsScript.Spreadsheet.DataValidationBuilder>(
        'DataValidationBuilder',
        ['requireCheckbox', 'setAllowInvalid', 'setHelpText', 'build'],
      );
    spreadsheetAppMock.newDataValidation.and.returnValue(
      dataValidatorBuilderMock,
    );
    dataValidatorBuilderMock.requireCheckbox.and.returnValue(
      dataValidatorBuilderMock,
    );
    dataValidatorBuilderMock.setAllowInvalid.and.returnValue(
      dataValidatorBuilderMock,
    );
    dataValidatorBuilderMock.setHelpText.and.returnValue(
      dataValidatorBuilderMock,
    );
    dataValidatorBuilderMock.build.and.returnValue(dataValidationMock);

    // Overwrite the global object with the mock.
    window.SpreadsheetApp = spreadsheetAppMock;
  });

  beforeEach(() => {
    spreadsheetMock =
      jasmine.createSpyObj<GoogleAppsScript.Spreadsheet.Spreadsheet>([
        'getNamedRanges',
        'getRangeByName',
      ]);

    sheetMock = jasmine.createSpyObj<GoogleAppsScript.Spreadsheet.Sheet>([
      'getName',
      'getRange',
      'getParent',
      'insertRows',
    ]);
    sheetMock.getName.and.returnValue('Sheet');
    sheetMock.getParent.and.returnValue(spreadsheetMock);
    sheetMock.getRange = jasmine
      .createSpy('getRange')
      .and.callFake(
        (row: number, column: number, numRows: number, numColumns: number) => {
          return createRangeMock(sheetMock, row, column, numRows, numColumns);
        },
      );
    sheetMock.insertRows.and.stub();

    sheetHandler = new SheetHandler(sheetMock);
  });

  describe('.getAdUnitId', () => {
    let adUnitIdRangeName: string;

    beforeEach(() => {
      // Configure the AD_UNIT_ID named range.
      adUnitIdRangeName = getLocalRangeName(
        /* sheet= */ sheetMock,
        /* rangeName= */ SheetHandler.NAMED_RANGE_AD_UNIT_ID,
      );
      rangeMock = createRangeMock(
        sheetMock,
        /* row= */ 1,
        /* column= */ 1,
        /* numRows= */ 1,
        /* numColumns= */ 1,
      );
      namedRangeMock = createNamedRangeMock(adUnitIdRangeName, rangeMock);
      spreadsheetMock.getNamedRanges.and.returnValue([namedRangeMock]);
      spreadsheetMock.getRangeByName
        .withArgs(adUnitIdRangeName)
        .and.returnValue(rangeMock);
    });

    it('returns the ad unit id', () => {
      rangeMock.getValue.and.returnValue(['5281981']);

      const adUnitId = sheetHandler.getAdUnitId();

      expect(adUnitId).toEqual('5281981');
    });

    it('throws an error if the AD_UNIT_ID range does not exist locally', () => {
      spreadsheetMock.getNamedRanges.and.returnValue([]);

      expect(() => {
        sheetHandler.getAdUnitId();
      }).toThrowError('AD_UNIT_ID range does not exist');
    });
  });

  describe('.getGoalType', () => {
    let goalTypeRangeName: string;

    beforeEach(() => {
      // Configure the GOAL_TYPE named range.
      goalTypeRangeName = getLocalRangeName(
        /* sheet= */ sheetMock,
        /* rangeName= */ SheetHandler.NAMED_RANGE_GOAL_TYPE,
      );
      rangeMock = createRangeMock(
        sheetMock,
        /* row= */ 1,
        /* column= */ 1,
        /* numRows= */ 1,
        /* numColumns= */ 1,
      );
      namedRangeMock = createNamedRangeMock(goalTypeRangeName, rangeMock);
      spreadsheetMock.getNamedRanges.and.returnValue([namedRangeMock]);
      spreadsheetMock.getRangeByName
        .withArgs(goalTypeRangeName)
        .and.returnValue(rangeMock);
    });

    it('returns the goal type', () => {
      rangeMock.getValue.and.returnValue('DAY');

      const goalType = sheetHandler.getGoalType();

      expect(goalType).toEqual(GoalType.DAY);
    });

    it('throws an error if the GOAL_TYPE range does not exist locally', () => {
      spreadsheetMock.getNamedRanges.and.returnValue([]);

      expect(() => {
        sheetHandler.getGoalType();
      }).toThrowError('GOAL_TYPE range does not exist');
    });
  });

  describe('.getScheduledEvents', () => {
    let scheduledEventsRangeName: string;

    beforeEach(() => {
      // Configure the SCHEDULED_EVENTS named range.
      scheduledEventsRangeName = getLocalRangeName(
        /* sheet= */ sheetMock,
        /* rangeName= */ SheetHandler.NAMED_RANGE_SCHEDULED_EVENTS,
      );
      rangeMock = createRangeMock(
        sheetMock,
        /* row= */ 1,
        /* column= */ 1,
        /* numRows= */ 5,
        /* numColumns= */ 4,
      );
      namedRangeMock = createNamedRangeMock(
        scheduledEventsRangeName,
        rangeMock,
      );
      spreadsheetMock.getNamedRanges.and.returnValue([namedRangeMock]);
      spreadsheetMock.getRangeByName
        .withArgs(scheduledEventsRangeName)
        .and.returnValue(rangeMock);
    });

    it('returns the scheduled events', () => {
      const scheduledEventRows = [
        [new Date('1/1/2024 20:00:00'), new Date('1/31/2024 02:00:00'), 10, ''],
        [new Date('2/1/2024 20:00:00'), new Date('2/31/2024 02:00:00'), 20, ''],
        ['', '', '', ''],
        ['', '', '', ''],
        ['', '', '', ''],
      ];
      rangeMock.getValues.and.returnValue(scheduledEventRows);

      const scheduledEvents = sheetHandler.getScheduledEvents();

      expect(scheduledEvents).toEqual([
        new ScheduledEvent('1/1/2024 20:00:00', '1/31/2024 02:00:00', 10, ''),
        new ScheduledEvent('2/1/2024 20:00:00', '2/31/2024 02:00:00', 20, ''),
      ]);
    });

    it('throws an error if the SCHEDULED_EVENTS range does not exist', () => {
      spreadsheetMock.getNamedRanges.and.returnValue([]);

      expect(() => {
        sheetHandler.getScheduledEvents();
      }).toThrowError('SCHEDULED_EVENTS range does not exist');
    });
  });

  describe('.getSelectedLineItems', () => {
    let lineItemsRangeName: string;

    beforeEach(() => {
      // Configure the LINE_ITEMS named range.
      lineItemsRangeName = getLocalRangeName(
        /* sheet= */ sheetMock,
        /* rangeName= */ SheetHandler.NAMED_RANGE_LINE_ITEMS,
      );
      rangeMock = createRangeMock(
        sheetMock,
        /* row= */ 1,
        /* column= */ 1,
        /* numRows= */ 10,
        /* numColumns= */ 6,
      );
      namedRangeMock = createNamedRangeMock(lineItemsRangeName, rangeMock);
      spreadsheetMock.getNamedRanges.and.returnValue([namedRangeMock]);
      spreadsheetMock.getRangeByName
        .withArgs(lineItemsRangeName)
        .and.returnValue(rangeMock);
    });

    it('returns no line items if none are selected', () => {
      const lineItemRows: any[][] = createFakeLineData(
        /* lineItemCount= */ 10,
        /* rowCount= */ 10,
      );
      rangeMock.getValues.and.returnValue(lineItemRows);

      const selectedLineItemRows = sheetHandler.getSelectedLineItems();

      expect(selectedLineItemRows.length).toBe(0);
    });

    it('returns only selected line items rows', () => {
      const lineItemRows: any[][] = createFakeLineData(
        /* lineItemCount= */ 10,
        /* rowCount= */ 10,
      );
      lineItemRows[0][0] = true; // Mark the first line item selected
      lineItemRows[1][0] = true; // Mark the second line item selected
      rangeMock.getValues.and.returnValue(lineItemRows);

      const selectedLineItemRows = sheetHandler.getSelectedLineItems();

      expect(selectedLineItemRows.length).toBe(2);
    });

    it('skips selected line item row with an invalid impression goal', () => {
      const lineItemRows: any[][] = createFakeLineData(
        /* lineItemCount= */ 1,
        /* rowCount= */ 10,
      );
      lineItemRows[0][0] = true; // Mark the first line item selected
      lineItemRows[0][5] = 'ABC'; // Set an invalid impression goal
      rangeMock.getValues.and.returnValue(lineItemRows);

      const selectedLineItemRows = sheetHandler.getSelectedLineItems();

      expect(selectedLineItemRows).toEqual([]);
    });

    it('skips selected line item row with an invalid line item id', () => {
      const lineItemRows: any[][] = createFakeLineData(
        /* lineItemCount= */ 1,
        /* rowCount= */ 10,
      );
      lineItemRows[0][0] = true; // Mark the first line item selected
      lineItemRows[0][1] = ''; // Set an invalid line item ID
      rangeMock.getValues.and.returnValue(lineItemRows);

      const selectedLineItemRows = sheetHandler.getSelectedLineItems();

      expect(selectedLineItemRows).toEqual([]);
    });

    it('throws an error if the LINE_ITEMS range does not exist', () => {
      spreadsheetMock.getNamedRanges.and.returnValue([]);

      expect(() => {
        sheetHandler.getSelectedLineItems();
      }).toThrowError('LINE_ITEMS range does not exist');
    });
  });

  describe('.writeLineItems', () => {
    let lineItemsRangeName: string;

    beforeEach(() => {
      // Configure the LINE_ITEMS named range.
      lineItemsRangeName = getLocalRangeName(
        /* sheet= */ sheetMock,
        /* rangeName= */ SheetHandler.NAMED_RANGE_LINE_ITEMS,
      );
      rangeMock = createRangeMock(
        sheetMock,
        /* row= */ 1,
        /* column= */ 1,
        /* numRows= */ 10,
        /* numColumns= */ 6,
      );
      namedRangeMock = createNamedRangeMock(lineItemsRangeName, rangeMock);
      spreadsheetMock.getNamedRanges.and.returnValue([namedRangeMock]);
      spreadsheetMock.getRangeByName
        .withArgs(lineItemsRangeName)
        .and.returnValue(rangeMock);
    });

    it('expands the named range (writing 30 rows to 10-row named range)', () => {
      const lineItemRows = Array(30).fill(MOCK_LINE_ITEM_ROW);

      sheetHandler.writeLineItems(lineItemRows);

      expect(namedRangeMock.setRange).toHaveBeenCalledTimes(1);
    });

    it('leave the named range unchanged (writing 5 rows to 10-row named range', () => {
      const lineItemRows = Array(5).fill(MOCK_LINE_ITEM_ROW);

      sheetHandler.writeLineItems(lineItemRows);

      expect(namedRangeMock.setRange).not.toHaveBeenCalled();
    });

    it('throws an error if the LINE_ITEMS range does not exist', () => {
      spreadsheetMock.getNamedRanges.and.returnValue([]);

      expect(() => {
        sheetHandler.writeLineItems([MOCK_LINE_ITEM_ROW]);
      }).toThrowError('LINE_ITEMS range does not exist');
    });
  });
});

/**
 * Returns a 2D array of fake line item values.
 * @param lineItemCount The number of rows to fill with line item data
 * @param rowCount The total number of rows to fill
 */
function createFakeLineData(lineItemCount: number, rowCount: number): any[][] {
  const lineItemRow = [
    /* selected= */ false,
    /* id= */ 12345678,
    /* name= */ 'Line Item Name',
    /* startDate= */ '2024-01-01 00:00:00',
    /* endDate= */ '2024-12-31 00:00:00',
    /* impressionGoal= */ 1000,
  ];
  const emptyRow = Array.from({length: lineItemRow.length}, () => '');

  if (lineItemCount < 0) {
    throw new RangeError('Line item count must be non-negative');
  }

  if (rowCount < 0) {
    throw new RangeError('Row count must be non-negative');
  }

  if (lineItemCount > rowCount) {
    throw new RangeError('Line item count must be less than row count');
  }

  const values = Array.from({length: lineItemCount}, () => [...lineItemRow]);

  values.push(...Array(rowCount - lineItemCount).fill(emptyRow));

  return values;
}

/** Returns a `NamedRange` mock with the given configuration. */
function createNamedRangeMock(
  name: string,
  range: GoogleAppsScript.Spreadsheet.Range,
): jasmine.SpyObj<GoogleAppsScript.Spreadsheet.NamedRange> {
  const namedRangeMock =
    jasmine.createSpyObj<GoogleAppsScript.Spreadsheet.NamedRange>([
      'getName',
      'getRange',
      'setRange',
    ]);

  namedRangeMock.getName.and.returnValue(name);
  namedRangeMock.getRange.and.returnValue(range);
  namedRangeMock.setRange.and.returnValue(namedRangeMock);

  return namedRangeMock;
}

/** Returns a `Range` mock with the given configuration. */
function createRangeMock(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  row: number,
  column: number,
  numRows: number,
  numColumns: number,
): jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Range> {
  const rangeMock = jasmine.createSpyObj<GoogleAppsScript.Spreadsheet.Range>([
    'getColumn',
    'getLastRow',
    'getNumColumns',
    'getNumRows',
    'getRow',
    'getSheet',
    'getValue',
    'getValues',
    'setDataValidation',
    'setValues',
  ]);

  rangeMock.getColumn.and.returnValue(column);
  rangeMock.getLastRow.and.returnValue(row + numRows - 1);
  rangeMock.getNumColumns.and.returnValue(numColumns);
  rangeMock.getNumRows.and.returnValue(numRows);
  rangeMock.getRow.and.returnValue(row);
  rangeMock.getSheet.and.returnValue(sheet);
  rangeMock.setValues.and.returnValue(rangeMock);

  return rangeMock;
}

/** Returns a local range name (e.g. 'Sheet1'!A1:B2). */
function getLocalRangeName(
  sheet: GoogleAppsScript.Spreadsheet.Sheet,
  rangeName: string,
): string {
  return `'${sheet.getName()}'!${rangeName}`;
}
