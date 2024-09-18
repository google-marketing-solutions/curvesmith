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

import {AdManagerClient} from 'gam_apps_script/ad_manager_client';
import {AdManagerService} from 'gam_apps_script/ad_manager_service';

import {
  AdManagerHandler,
  LineItemDtoPage,
  LineItemFilter,
} from './ad_manager_handler';
import * as ad_manager from './typings/ad_manager';

describe('AdManagerHandler', () => {
  let mockClient: jasmine.SpyObj<AdManagerClient>;
  let mockInventoryService: jasmine.SpyObj<AdManagerService>;
  let mockNetwork: jasmine.SpyObj<ad_manager.Network>;
  let mockNetworkService: jasmine.SpyObj<AdManagerService>;
  let mockPqlService: jasmine.SpyObj<AdManagerService>;

  beforeEach(() => {
    // Initializes fake client services for testing
    mockInventoryService = jasmine.createSpyObj('AdManagerService', [
      'performOperation',
    ]);

    mockNetwork = jasmine.createSpyObj('Network', ['timeZone']);

    Object.defineProperty(mockNetwork, 'timeZone', {
      get: () => 'America/Los_Angeles',
    });

    mockNetworkService = jasmine.createSpyObj('AdManagerService', [
      'performOperation',
    ]);
    mockNetworkService.performOperation
      .withArgs('getCurrentNetwork')
      .and.returnValue(mockNetwork);

    mockPqlService = jasmine.createSpyObj('AdManagerService', [
      'performOperation',
    ]);

    mockClient = jasmine.createSpyObj('AdManagerClient', ['getService']);
    mockClient.getService
      .withArgs('InventoryService')
      .and.returnValue(mockInventoryService);
    mockClient.getService
      .withArgs('NetworkService')
      .and.returnValue(mockNetworkService);
    mockClient.getService
      .withArgs('PublisherQueryLanguageService')
      .and.returnValue(mockPqlService);
  });

  describe('.getAdUnitIds', () => {
    it('returns an empty array if a parent ad unit ID is not provided', () => {
      const adUnitId = '';
      const adManagerHandler = new AdManagerHandler(mockClient);

      expect(adManagerHandler.getAdUnitIds(adUnitId)).toEqual([]);
    });

    it('returns the parent ad unit ID even if no children are found', () => {
      const adUnitId = '1234';
      const adManagerHandler = new AdManagerHandler(mockClient);
      mockInventoryService.performOperation.and.returnValue({results: []});

      expect(adManagerHandler.getAdUnitIds(adUnitId)).toEqual([adUnitId]);
    });
  });

  describe('.getDateTime', () => {
    it('returns a DateTime that matches the Date', () => {
      // Forces the creation of a specific UTC date that should align with the
      // expected result. Normally the Date constructor will not be provided an
      // explicit timezone, instead defaulting to the local one which should
      // match Ad Manager. We simulate this behavior with the GMT offset.
      const date = new Date('2024-01-01 GMT-08:00');
      const adManagerHandler = new AdManagerHandler(mockClient);

      expect(adManagerHandler.getDateTime(date)).toEqual({
        date: {
          year: 2024,
          month: 1,
          day: 1,
        },
        hour: 0,
        minute: 0,
        second: 0,
        timeZoneId: 'America/Los_Angeles',
      });
    });
  });

  describe('.getDateString', () => {
    it('returns a timezone agnostic string for the given date', () => {
      const dateTime: ad_manager.DateTime = {
        date: {
          year: 2024,
          month: 5,
          day: 28,
        },
        hour: 2,
        minute: 57,
        second: 0,
        timeZoneId: 'America/Los_Angeles',
      };
      const adManagerHandler = new AdManagerHandler(mockClient);

      expect(adManagerHandler.getDateString(dateTime)).toEqual(
        '2024-05-28 02:57:00',
      );
    });
  });

  describe('.getLineItemDtoPage', () => {
    const columnTypes: ad_manager.ColumnType[] = [
      {labelName: 'id'},
      {labelName: 'name'},
      {labelName: 'startDateTime'},
      {labelName: 'endDateTime'},
      {labelName: 'targeting'},
      {labelName: 'unitsBought'},
    ];

    // Our focus here will be testing any post-processing of the results, not
    // the actual Ad Manager API calls, so we'll always return the same fake
    // data and validate that the results are filtered as expected.
    const row1AdUnit1234 = createFakeRowForResultSet(
      /* id= */ 1,
      /* name= */ 'Line Item 1',
      /* endDate= */ {
        year: 2025,
        month: 1,
        day: 1,
      },
      /* adUnitId = */ '1234',
    );
    const row2AdUnit1234 = createFakeRowForResultSet(
      /* id= */ 2,
      /* name= */ 'Line Item 2 (ROS)',
      /* endDate= */ {
        year: 2025,
        month: 1,
        day: 1,
      },
      /* adUnitId = */ '1234',
    );
    const row3AdUnit5678 = createFakeRowForResultSet(
      /* id= */ 3,
      /* name= */ 'Line Item 3 - ROS',
      /* endDate= */ {
        year: 2025,
        month: 1,
        day: 1,
      },
      /* adUnitId = */ '5678',
    );
    // Simulate line item returned because of grace period.
    const row4AdUnit1234Grace = createFakeRowForResultSet(
      /* id= */ 4,
      /* name= */ 'Line Item 4 - Grace Period',
      /* endDate= */ {
        year: 2024,
        month: 31,
        day: 12,
      },
      /* adUnitId = */ '1234',
    );

    beforeEach(() => {
      const lineItemDtoPage: ad_manager.ResultSet = {
        columnTypes,
        rows: [
          row1AdUnit1234,
          row2AdUnit1234,
          row3AdUnit5678,
          row4AdUnit1234Grace,
        ],
      };

      mockPqlService.performOperation.and.returnValue(lineItemDtoPage);
    });

    it('cache the PqlService upon first call', () => {
      const lineItemFilter: LineItemFilter = {
        adUnitIds: ['1234'],
        latestStartDate: new Date('2024-01-01 00:00:00'),
        earliestEndDate: new Date('2025-01-01 00:00:00'),
        nameFilter: '',
      };
      const adManagerHandler = new AdManagerHandler(mockClient);

      // The first call should cache the LineItemService.
      adManagerHandler.getLineItemDtoPage(lineItemFilter, 0, 1);
      // The second call should use the cached LineItemService.
      adManagerHandler.getLineItemDtoPage(lineItemFilter, 1, 1);

      const serviceNames = mockClient.getService.calls
        .allArgs()
        .map((args) => args[0]);
      expect(
        serviceNames.filter((x) => x === 'PublisherQueryLanguageService')
          .length,
      ).toBe(1);
    });

    it('filters line items with end dates that are too early', () => {
      const lineItemDtoPage = getLineItemsByAdUnitFilter(mockClient, []);

      // Row 4 should be filtered out because its end date is too early.
      expect(lineItemDtoPage.values).toEqual([
        jasmine.objectContaining({id: row1AdUnit1234.values[0].value}),
        jasmine.objectContaining({id: row2AdUnit1234.values[0].value}),
        jasmine.objectContaining({id: row3AdUnit5678.values[0].value}),
      ]);
    });

    it('returns an empty array if no line items match', () => {
      const adUnitIds = ['9999'];

      const lineItemDtoPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      expect(lineItemDtoPage.values).toEqual([]);
    });

    it('returns the full results if no ad unit ID is provided', () => {
      const adUnitIds: string[] = [];

      const lineItemDtoPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      // Row 4 should be filtered out because its end date is too early.
      expect(lineItemDtoPage.values).toEqual([
        jasmine.objectContaining({id: row1AdUnit1234.values[0].value}),
        jasmine.objectContaining({id: row2AdUnit1234.values[0].value}),
        jasmine.objectContaining({id: row3AdUnit5678.values[0].value}),
      ]);
    });

    it('returns line items that match any of the provided ad unit IDs', () => {
      const adUnitIds = ['1234', '5678'];

      const lineItemDtoPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      // Row 4 should be filtered out because its end date is too early.
      expect(lineItemDtoPage.values).toEqual([
        jasmine.objectContaining({id: row1AdUnit1234.values[0].value}),
        jasmine.objectContaining({id: row2AdUnit1234.values[0].value}),
        jasmine.objectContaining({id: row3AdUnit5678.values[0].value}),
      ]);
    });

    it('returns line items that match the provided ad unit ID', () => {
      const adUnitIds = ['1234'];

      const lineItemDtoPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      // Row 3 should be filtered out because it targets a different ad unit.
      // Row 4 should be filtered out because its end date is too early.
      expect(lineItemDtoPage.values).toEqual([
        jasmine.objectContaining({id: row1AdUnit1234.values[0].value}),
        jasmine.objectContaining({id: row2AdUnit1234.values[0].value}),
      ]);
    });

    it('returns line items that match the provided name filter', () => {
      const lineItemFilter: LineItemFilter = {
        adUnitIds: [],
        latestStartDate: new Date('2024-01-01 00:00:00'),
        earliestEndDate: new Date('2025-01-01 00:00:00'),
        nameFilter: 'ROS',
      };
      const adManagerHandler = new AdManagerHandler(mockClient);

      const lineItemDtoPage = adManagerHandler.getLineItemDtoPage(
        lineItemFilter,
        0,
        100,
      );

      expect(lineItemDtoPage.values).toEqual([
        jasmine.objectContaining({name: row2AdUnit1234.values[1].value}),
        jasmine.objectContaining({name: row3AdUnit5678.values[1].value}),
      ]);
    });
  });
});

function createFakeRowForResultSet(
  id: number,
  name: string,
  endDate: ad_manager.Date,
  adUnitId: string,
): ad_manager.Row {
  return {
    values: [
      {/* id= */ value: id},
      {/* name= */ value: name},
      {
        /* startDateTime= */ value: {
          date: {
            year: 2024,
            month: 1,
            day: 1,
          },
          hour: 0,
          minute: 0,
          second: 0,
          timeZoneId: 'America/Los_Angeles',
        },
      },
      {
        /* endDateTime= */ value: {
          date: endDate,
          hour: 0,
          minute: 0,
          second: 0,
          timeZoneId: 'America/Los_Angeles',
        },
      },
      {
        /* targeting= */ value: {
          inventoryTargeting: {
            targetedAdUnits: [{adUnitId: adUnitId}],
          },
        },
      },
      {
        /* unitsBought= */ value: 1000,
      },
    ],
  };
}

/** Helper function to get the results of a call to `getLineItemDtoPage`. */
function getLineItemsByAdUnitFilter(
  client: AdManagerClient,
  adUnitIds: string[],
): LineItemDtoPage {
  const adManagerHandler = new AdManagerHandler(client);

  // Both dates are irrelevant because they are PQL arguments and we are mocking
  // the LineItemService call.
  const lineItemFilter: LineItemFilter = {
    adUnitIds,
    latestStartDate: new Date('2024-01-01 00:00:00'),
    earliestEndDate: new Date('2025-01-01 00:00:00'),
    nameFilter: '',
  };

  return adManagerHandler.getLineItemDtoPage(lineItemFilter, 0, 100);
}
