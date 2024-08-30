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

import {AdManagerHandler, LineItemFilter} from './ad_manager_handler';
import * as ad_manager from './typings/ad_manager';

describe('AdManagerHandler', () => {
  let mockClient: jasmine.SpyObj<AdManagerClient>;
  let mockInventoryService: jasmine.SpyObj<AdManagerService>;
  let mockLineItemService: jasmine.SpyObj<AdManagerService>;
  let mockNetwork: jasmine.SpyObj<ad_manager.Network>;
  let mockNetworkService: jasmine.SpyObj<AdManagerService>;

  beforeEach(() => {
    // Initializes fake client services for testing
    mockInventoryService = jasmine.createSpyObj('AdManagerService', [
      'performOperation',
    ]);

    mockLineItemService = jasmine.createSpyObj('AdManagerService', [
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

    mockClient = jasmine.createSpyObj('AdManagerClient', ['getService']);
    mockClient.getService
      .withArgs('InventoryService')
      .and.returnValue(mockInventoryService);
    mockClient.getService
      .withArgs('LineItemService')
      .and.returnValue(mockLineItemService);
    mockClient.getService
      .withArgs('NetworkService')
      .and.returnValue(mockNetworkService);
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

  describe('.getLineItemsByFilter', () => {
    let mockLine1AdUnit1234: jasmine.SpyObj<ad_manager.LineItem>;
    let mockLine2AdUnit1234: jasmine.SpyObj<ad_manager.LineItem>;
    let mockLineAdUnit5678: jasmine.SpyObj<ad_manager.LineItem>;

    beforeEach(() => {
      // Our focus here will be testing any post-processing of the results, not
      // the actual Ad Manager API calls, so we'll always return the same fake
      // data and validate that the results are filtered as expected.
      mockLine1AdUnit1234 = jasmine.createSpyObj('LineItem', [], {
        endDateTime: {
          date: {
            year: 2026,
            month: 1,
            day: 1,
          },
          hour: 0,
          minute: 0,
          second: 0,
          timeZoneId: 'America/Los_Angeles',
        },
        targeting: {
          inventoryTargeting: {
            targetedAdUnits: [{adUnitId: '1234'}],
          },
        },
      });
      mockLine2AdUnit1234 = jasmine.createSpyObj('LineItem', [], {
        endDateTime: {
          date: {
            year: 2026,
            month: 1,
            day: 1,
          },
          hour: 0,
          minute: 0,
          second: 0,
          timeZoneId: 'America/Los_Angeles',
        },
        targeting: {
          inventoryTargeting: {
            targetedAdUnits: [{adUnitId: '1234'}],
          },
        },
      });
      mockLineAdUnit5678 = jasmine.createSpyObj('LineItem', [], {
        endDateTime: {
          date: {
            year: 2026,
            month: 1,
            day: 1,
          },
          hour: 0,
          minute: 0,
          second: 0,
          timeZoneId: 'America/Los_Angeles',
        },
        targeting: {
          inventoryTargeting: {
            targetedAdUnits: [{adUnitId: '5678'}],
          },
        },
      });

      const lineItemPageFake: ad_manager.LineItemPage = {
        results: [mockLine1AdUnit1234, mockLine2AdUnit1234, mockLineAdUnit5678],
        startIndex: 0,
        totalResultSetSize: 3,
      };

      mockLineItemService.performOperation.and.returnValue(lineItemPageFake);
    });

    it('cache the LineItemService upon first call', () => {
      const lineItemFilter: LineItemFilter = {
        adUnitIds: ['1234'],
        latestStartDate: new Date('2024-01-01 00:00:00'),
        earliestEndDate: new Date('2025-01-01 00:00:00'),
      };
      const adManagerHandler = new AdManagerHandler(mockClient);

      // The first call should cache the LineItemService.
      adManagerHandler.getLineItemsByFilter(lineItemFilter, 0);
      // The second call should use the cached LineItemService.
      adManagerHandler.getLineItemsByFilter(lineItemFilter, 1);

      const serviceNames = mockClient.getService.calls.allArgs().map(
        (args) => args[0],
      );
      expect(serviceNames.filter(x => x === 'LineItemService').length).toBe(1);
    });

    it('filters line items with end dates that are too early', () => {
      // Simulate line item returned because of grace period.
      mockLineAdUnit5678.endDateTime.date = {
        year: 2024,
        month: 12,
        day: 31,
      };
      const adUnitIds: string[] = [];

      const lineItemPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      expect(lineItemPage.results).toEqual([
        mockLine1AdUnit1234,
        mockLine2AdUnit1234,
      ]);
    });

    it('returns an empty array if no line items match', () => {
      const adUnitIds = ['9999'];

      const lineItemPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      expect(lineItemPage.results).toEqual([]);
    });

    it('returns the full results if no ad unit ID is provided', () => {
      const adUnitIds: string[] = [];

      const lineItemPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      expect(lineItemPage.results).toEqual([
        mockLine1AdUnit1234,
        mockLine2AdUnit1234,
        mockLineAdUnit5678,
      ]);
    });

    it('returns line items that match any of the provided ad unit IDs', () => {
      const adUnitIds = ['1234', '5678'];

      const lineItemPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      expect(lineItemPage.results).toEqual([
        mockLine1AdUnit1234,
        mockLine2AdUnit1234,
        mockLineAdUnit5678,
      ]);
    });

    it('returns line items that match the provided ad unit ID', () => {
      const adUnitIds = ['1234'];

      const lineItemPage = getLineItemsByAdUnitFilter(mockClient, adUnitIds);

      expect(lineItemPage.results).toEqual([
        mockLine1AdUnit1234,
        mockLine2AdUnit1234,
      ]);
    });
  });
});

/** Helper function to get the results of a call to `getLineItemsByFilter`. */
function getLineItemsByAdUnitFilter(
  client: AdManagerClient,
  adUnitIds: string[],
): ad_manager.LineItemPage {
  const adManagerHandler = new AdManagerHandler(client);

  // Both dates are irrelevant because they are PQL arguments and we are mocking
  // the LineItemService call.
  const lineItemFilter: LineItemFilter = {
    adUnitIds,
    latestStartDate: new Date('2024-01-01 00:00:00'),
    earliestEndDate: new Date('2025-01-01 00:00:00'),
  };

  return adManagerHandler.getLineItemsByFilter(lineItemFilter, 0);
}
