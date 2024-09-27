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

import {
  AdManagerHandler,
  LineItemDto,
  LineItemDtoPage,
} from './ad_manager_handler';
import {TEST_ONLY} from './curvesmith';
import {CurveTemplate, GoalType, ScheduledEvent} from './custom_curve';
import {SheetHandler, SpreadsheetHandler} from './sheet_handler';

const {
  copyTemplate,
  getAdManagerHandler,
  beginLoadLineItems,
  getLineItemPreviews,
  getTaskProgress,
  loadLineItems,
  setTaskProgress,
} = TEST_ONLY;

describe('curvesmith', () => {
  const propertiesFake = new Map<string, string>();
  let baseUiMock: jasmine.SpyObj<GoogleAppsScript.Base.Ui>;
  let spreadsheetAppMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.SpreadsheetApp>;
  let propertiesServiceMock: jasmine.SpyObj<GoogleAppsScript.Properties.PropertiesService>;
  let propertiesMock: jasmine.SpyObj<GoogleAppsScript.Properties.Properties>;

  beforeAll(() => {
    // Outside of the Google Apps Script environment, PropertiesService is not
    // available. This creates a mock that uses a Map as a backing store.
    propertiesMock = jasmine.createSpyObj('Properties', [
      'deleteProperty',
      'getKeys',
      'getProperty',
      'setProperty',
    ]);
    propertiesMock.deleteProperty.and.callFake((key: string) => {
      propertiesFake.delete(key);
      return propertiesMock;
    });
    propertiesMock.getKeys.and.callFake(() => {
      return [...propertiesFake.keys()];
    });
    propertiesMock.getProperty.and.callFake((key: string) => {
      return propertiesFake.get(key) ?? null;
    });
    propertiesMock.setProperty.and.callFake((key: string, value: string) => {
      propertiesFake.set(key, value);
      return propertiesMock;
    });
    propertiesServiceMock = jasmine.createSpyObj('PropertiesService', [
      'getScriptProperties',
      'getUserProperties',
    ]);
    propertiesServiceMock.getScriptProperties.and.returnValue(propertiesMock);
    propertiesServiceMock.getUserProperties.and.returnValue(propertiesMock);

    // Similarly, SpreadsheetApp is unavailable outside of the Apps Script
    // environment. This creates a mock that uses a fake UI implementation.
    baseUiMock = jasmine.createSpyObj('Ui', ['alert']);
    Object.defineProperty(baseUiMock, 'ButtonSet', {
      value: {YES_NO: 'YES_NO'},
      writable: false,
    });
    Object.defineProperty(baseUiMock, 'Button', {
      value: {NO: 'NO'},
      writable: false,
    });
    spreadsheetAppMock = jasmine.createSpyObj('SpreadsheetApp', ['getUi']);
    spreadsheetAppMock.getUi.and.returnValue(baseUiMock);

    // Overwrite the global objects with the mocks.
    window.SpreadsheetApp = spreadsheetAppMock;
    window.PropertiesService = propertiesServiceMock;
  });

  beforeEach(() => {
    propertiesFake.clear();
  });

  describe('copyTemplate', () => {
    let rangeMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Range>;
    let spreadsheetMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Spreadsheet>;
    let spreadsheetHandler: SpreadsheetHandler;

    beforeEach(() => {
      rangeMock = jasmine.createSpyObj('Range', ['getValue']);
      spreadsheetMock = jasmine.createSpyObj('Spreadsheet', [
        'getRangeByName',
        'getSheetByName',
      ]);
      spreadsheetMock.getRangeByName.and.returnValue(null);

      spreadsheetHandler = new SpreadsheetHandler(spreadsheetMock);
    });

    it('copies the template sheet', () => {
      const templateSheetName = 'Template';
      const templateSheetMock = jasmine.createSpyObj('Sheet', [
        'activate',
        'copyTo',
      ]);
      templateSheetMock.activate.and.returnValue(templateSheetMock);
      templateSheetMock.copyTo.and.returnValue(templateSheetMock);
      rangeMock.getValue.and.returnValue(templateSheetName);
      spreadsheetMock.getRangeByName
        .withArgs(SpreadsheetHandler.NAMED_RANGE_TEMPLATE_SHEET_NAME)
        .and.returnValue(rangeMock);
      spreadsheetMock.getSheetByName
        .withArgs(templateSheetName)
        .and.returnValue(templateSheetMock);

      copyTemplate(spreadsheetHandler);

      expect(templateSheetMock.copyTo).toHaveBeenCalled();
    });

    it('throws an error if the template sheet is missing', () => {
      const templateSheetName = 'Template';
      rangeMock.getValue.and.returnValue(templateSheetName);
      spreadsheetMock.getRangeByName
        .withArgs(SpreadsheetHandler.NAMED_RANGE_TEMPLATE_SHEET_NAME)
        .and.returnValue(rangeMock);
      spreadsheetMock.getSheetByName
        .withArgs(templateSheetName)
        .and.returnValue(null);

      expect(() => {
        copyTemplate(spreadsheetHandler);
      }).toThrowError('Template sheet does not exist');
    });

    it('throws an error if the TEMPLATE_SHEET_NAME range does not exist', () => {
      expect(() => {
        copyTemplate(spreadsheetHandler);
      }).toThrowError('TEMPLATE_SHEET_NAME range does not exist');
    });
  });

  describe('getAdManagerHandler', () => {
    let rangeMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Range>;
    let spreadsheetMock: jasmine.SpyObj<GoogleAppsScript.Spreadsheet.Spreadsheet>;
    let spreadsheetHandler: SpreadsheetHandler;

    beforeEach(() => {
      rangeMock = jasmine.createSpyObj('Range', ['getValue']);
      spreadsheetMock = jasmine.createSpyObj('Spreadsheet', ['getRangeByName']);
      spreadsheetMock.getRangeByName.and.returnValue(null);

      spreadsheetHandler = new SpreadsheetHandler(spreadsheetMock);
    });

    it('throws an error if the API_VERSION range does not exist', () => {
      spreadsheetMock.getRangeByName
        .withArgs(SpreadsheetHandler.NAMED_RANGE_AD_MANAGER_NETWORK_ID)
        .and.returnValue(rangeMock);

      expect(() => {
        getAdManagerHandler(spreadsheetHandler);
      }).toThrowError('API_VERSION range does not exist');
    });

    it('throws an error if the NETWORK_ID range does not exist', () => {
      spreadsheetMock.getRangeByName
        .withArgs(SpreadsheetHandler.NAMED_RANGE_AD_MANAGER_API_VERSION)
        .and.returnValue(rangeMock);

      expect(() => {
        getAdManagerHandler(spreadsheetHandler);
      }).toThrowError('NETWORK_ID range does not exist');
    });
  });

  describe('getLineItemPreviews', () => {
    it('returns expected preview for the selected line item', () => {
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 80, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);
      const sheetHandlerMock = jasmine.createSpyObj('SheetHandler', [
        'getCurveTemplate',
        'getSelectedLineItems',
      ]);
      sheetHandlerMock.getCurveTemplate.and.returnValue(curveTemplate);
      sheetHandlerMock.getSelectedLineItems.and.returnValue([
        {
          selected: true,
          id: 12345678,
          name: 'Line Item',
          startDate: '3/27/2024',
          endDate: '4/1/2024 23:59:00',
          impressionGoal: 500,
        },
      ]);

      const lineItemPreviews = getLineItemPreviews(sheetHandlerMock);

      expect(lineItemPreviews).toEqual([
        jasmine.objectContaining({
          curveGoals: [
            {
              description: 'Pre-Event [A]',
              startDate: new Date('3/27/2024 00:00:00').toISOString(),
              goalPercent: 12.561679905035465,
              impressionGoal: 62.80839952517732,
            },
            {
              description: 'A',
              startDate: new Date('3/27/2024 20:00:00').toISOString(),
              goalPercent: 13.334876721842805,
              impressionGoal: 66.67438360921403,
            },
            {
              description: 'Post-Events',
              startDate: new Date('3/28/2024 02:00:00').toISOString(),
              goalPercent: 74.10344337312172,
              impressionGoal: 370.51721686560865,
            },
          ],
        }),
      ]);
    });

    it('throws an error if no line items are selected', () => {
      const events = [new ScheduledEvent('3/27/2024', '3/28/2024', 80, '')];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);
      const sheetHandlerMock = jasmine.createSpyObj('SheetHandler', [
        'getCurveTemplate',
        'getSelectedLineItems',
      ]);
      sheetHandlerMock.getCurveTemplate.and.returnValue(curveTemplate);
      sheetHandlerMock.getSelectedLineItems.and.returnValue([]);

      expect(() => getLineItemPreviews(sheetHandlerMock)).toThrowError(
        'No line items are selected',
      );
    });

    it('throws an error if a valid template is not present', () => {
      expect(() => getLineItemPreviews()).toThrowError();
    });
  });

  describe('loadLineItems', () => {
    let adManagerHandlerMock: jasmine.SpyObj<AdManagerHandler>;
    let sheetHandlerMock: jasmine.SpyObj<SheetHandler>;

    beforeEach(() => {
      adManagerHandlerMock = jasmine.createSpyObj('AdManagerHandler', [
        'getAdUnitIds',
        'getDateString',
        'getLineItemCount',
        'getLineItemDtoPage',
      ]);
      adManagerHandlerMock.getAdUnitIds.and.returnValue(['1234', '5678']);
      adManagerHandlerMock.getLineItemCount.and.returnValue(1);

      sheetHandlerMock = jasmine.createSpyObj('SheetHandler', [
        'clearLineItems',
        'getAdUnitId',
        'getNameFilter',
        'getScheduledEvents',
        'writeLineItems',
      ]);
      sheetHandlerMock.getAdUnitId.and.returnValue('1234');
      sheetHandlerMock.getNameFilter.and.returnValue('');
      sheetHandlerMock.getScheduledEvents.and.returnValue([
        new ScheduledEvent('1/1/2024', '1/2/2024', 33, ''),
        new ScheduledEvent('1/3/2024', '1/4/2024', 33, ''),
      ]);
    });

    it('clears any existing line item metadata', () => {
      const lineItemDtoPage: LineItemDtoPage = {
        values: [],
        endOfResults: true,
      };
      adManagerHandlerMock.getLineItemDtoPage.and.returnValue(lineItemDtoPage);

      beginLoadLineItems(adManagerHandlerMock, sheetHandlerMock);

      expect(sheetHandlerMock.clearLineItems).toHaveBeenCalled();
    });

    it('sets task progress to 100% upon completion', () => {
      const lineItemDtoPage: LineItemDtoPage = {
        values: createLineItemDtos(11),
        endOfResults: true,
      };
      adManagerHandlerMock.getDateString.and.returnValue('2024-01-01 00:00:00');
      adManagerHandlerMock.getLineItemDtoPage.and.returnValue(lineItemDtoPage);
      setTaskProgress('Fake', 0, lineItemDtoPage.values.length);

      loadLineItems(0, 50, adManagerHandlerMock, sheetHandlerMock);

      const taskProgress = getTaskProgress();
      expect(taskProgress.current / taskProgress.total).toBe(1);
    });

    it('throws an error if no scheduled events are specified', () => {
      sheetHandlerMock.getScheduledEvents.and.returnValue([]);

      expect(() => {
        beginLoadLineItems(adManagerHandlerMock, sheetHandlerMock);
      }).toThrowError('No scheduled events are specified');
    });

    it('writes line item metadata to the sheet', () => {
      const lineItemDtoPage: LineItemDtoPage = {
        values: createLineItemDtos(2),
        endOfResults: true,
      };
      adManagerHandlerMock.getDateString.and.returnValue('2024-01-01 00:00:00');
      adManagerHandlerMock.getLineItemDtoPage.and.returnValue(lineItemDtoPage);
      setTaskProgress('Fake', 0, lineItemDtoPage.values.length);

      loadLineItems(0, 50, adManagerHandlerMock, sheetHandlerMock);

      // Will write to the sheet because all pages have been loaded
      expect(sheetHandlerMock.writeLineItems).toHaveBeenCalledWith([
        jasmine.objectContaining({id: 1, name: 'mock-line-item-1'}),
        jasmine.objectContaining({id: 2, name: 'mock-line-item-2'}),
      ]);
    });
  });
});

/**
 * Generates an array of mock LineItem objects for tests.
 *
 * Each line item has a unique ID, a name formatted as "mock-line-item-{id}",
 * and basic placeholders for remaining fields.
 * @param count The number of mock line items to create.
 * @returns An array of mock LineItem objects.
 */
function createLineItemDtos(count: number): LineItemDto[] {
  return Array.from({length: count}, (_, index) => {
    const lineItemId = index + 1;
    return {
      id: lineItemId,
      name: `mock-line-item-${lineItemId}`,
      startDate: '2024-01-01 00:00:00',
      endDate: '2025-01-01 00:00:00',
      impressionGoal: 1000,
    };
  });
}
