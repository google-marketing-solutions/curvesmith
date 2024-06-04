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
 * @fileoverview Uses the GAM Apps Script library and OAuth credentials provided
 * by the current Apps Script environment (e.g. Google Sheets) to interact with
 * a Google Ad Manager network.
 */

import {CurveTemplate, FlightDetails} from './custom_curve';
import * as ad_manager from './typings/ad_manager';

import {AdManagerClient} from 'gam_apps_script/ad_manager_client';
import {AdManagerService} from 'gam_apps_script/ad_manager_service';
import {StatementBuilder} from 'gam_apps_script/statement_builder';

/**
 * A collection of configuration values that will be used to filter line items
 * when retrieving them from Ad Manager. The goal is to only populate the active
 * sheet with lines that are capable of supporting the curve template.
 */
export interface LineItemFilter {
  /** Line items must target at least one of these ad unit IDs. */
  adUnitIds: string[];

  /** Line items must begin their flights no later than this date. */
  latestStartDate: Date;

  /** Line items must complete their flights no earlier than this date. */
  earliestEndDate: Date;
}

/**
 * Returns a new Ad Manager client using the OAuth token from the current
 * active user.
 * @param networkId The Ad Manager network ID
 * @param apiVersion The Ad Manager API version
 */
export function createAdManagerClient(
  networkId: string,
  apiVersion: string,
): AdManagerClient {
  return new AdManagerClient(
    ScriptApp.getOAuthToken(),
    'curvesmith',
    networkId,
    apiVersion,
  );
}

/**
 * The required sum of all custom pacing goals in a delivery curve. Ad Manager
 * will reject the curve if this value is not met exactly.
 */
const TOTAL_MILLIPERCENT_REQUIRED = 100000;

/**
 * A class for interacting with a Google Ad Manager network to retrieve line
 * items and, when appropriate, update them with custom delivery curves.
 *
 * Ad Manager service endpoints are initialized upon first use and then cached
 * afterward in `serviceCache` in order to improve performance.
 */
export class AdManagerHandler {
  private readonly serviceCache: Map<string, AdManagerService>;

  constructor(readonly client: AdManagerClient) {
    this.serviceCache = new Map<string, AdManagerService>();
  }

  /**
   * The number of objects to request at a time. This value was empirically
   * determined to be the optimal tradeoff between UX and performance.
   */
  static readonly AD_MANAGER_API_PAGE_LIMIT = 10;

  /**
   * Updates each of the provided line items with a custom delivery curve. This
   * is a local change and the line items will not be updated in Ad Manager
   * until `uploadLineItems` is called.
   * @param lineItems An array of line items to update
   * @param template The curve template to use when generating the curve
   */
  applyCurveToLineItems(
    lineItems: ad_manager.LineItem[],
    template: CurveTemplate,
  ) {
    for (const lineItem of lineItems) {
      const flight = new FlightDetails(
        /** start= */ this.getDateString(lineItem.startDateTime),
        /** end= */ this.getDateString(lineItem.endDateTime),
        /** goal= */ lineItem.primaryGoal.units,
      );
      const curveSegments = template.generateCurveSegments(flight);

      let totalGoalMillipercent = 0;
      const customPacingGoals: ad_manager.CustomPacingGoal[] = [];

      for (const segment of curveSegments) {
        const goalMillipercent = Math.round(segment.goalPercent * 1000);

        customPacingGoals.push({
          startDateTime: this.getDateTime(segment.start),
          useLineItemStartDateTime: segment.start === flight.start,
          amount: goalMillipercent,
        });

        totalGoalMillipercent += goalMillipercent;
      }

      const difference = TOTAL_MILLIPERCENT_REQUIRED - totalGoalMillipercent;

      // Account for any precision errors by adding the difference
      customPacingGoals[curveSegments.length - 1].amount += difference;

      lineItem.deliveryForecastSource = 'CUSTOM_PACING_CURVE';
      lineItem.customPacingCurve = {
        customPacingGoalUnit: 'MILLI_PERCENT',
        customPacingGoals: customPacingGoals,
      };
    }
  }

  /**
   * Sets the delivery forecast source of the provided line items to historical.
   * This is a local change and the line items will not be updated in Ad Manager
   * until `uploadLineItems` is called.
   * @param lineItems An array of line items to update
   */
  applyHistoricalToLineItems(lineItems: ad_manager.LineItem[]) {
    for (const lineItem of lineItems) {
      lineItem.deliveryForecastSource = 'HISTORICAL';
    }
  }

  /**
   * Returns an array of all ad unit IDs starting from the provided ad unit ID
   * and traversing down the tree.
   */
  getAdUnitIds(adUnitId: string): string[] {
    if (!adUnitId?.trim()) {
      return [];
    }

    const adUnitIds: string[] = [adUnitId];

    const statement = new StatementBuilder()
      .where('parentId = :parentId AND status = :status')
      .withBindVariable('parentId', adUnitId)
      .withBindVariable('status', 'ACTIVE')
      .toStatement();

    const inventoryService = this.getService('InventoryService');
    const adUnitPage = inventoryService.performOperation(
      'getAdUnitsByStatement',
      statement,
    ) as ad_manager.AdUnitPage;

    for (const adUnit of adUnitPage.results) {
      if (adUnit.explicitlyTargeted) {
        continue;
      }

      if (adUnit.hasChildren) {
        adUnitIds.push(...this.getAdUnitIds(adUnit.id));
      } else {
        adUnitIds.push(adUnit.id);
      }
    }

    return adUnitIds;
  }

  /**
   * Returns a representative `DateTime` object for the provided date.
   *
   * While internally Date objects are stored as an offset from the UTC epoch,
   * parsing a date string without a time zone offset will initialize the Date
   * object relative to the local time zone. For example, if the spreadsheet's
   * time zone is "America/New_York", then creating a Date object from the value
   * "1/1/2024 10:00" will effectively be initialized as "1/1/2024 10:00 GMT-5".
   *
   * Consequently, this function relies on two assumptions:
   *  1. The spreadsheet's time zone matches the Ad Manager network time zone.
   *  2. The provided date has been initialized in the given time zone.
   *
   * Otherwise, an unexpected result is likely to be returned.
   * @param date A date value initialized using the same time zone as Ad Manager
   */
  getDateTime(date: Date): ad_manager.DateTime {
    const timeZoneId = this.getTimeZoneId();
    const localeString = date.toLocaleString('en-US', {
      hour12: false,
      timeZone: timeZoneId,
    });

    const dateMatch = localeString.match(
      /(\d+)\/(\d+)\/(\d+), (\d+):(\d+):(\d+)/,
    );

    if (!dateMatch) {
      throw new Error('An unexpected error occurred while parsing the date');
    }

    let [, month, day, year, hour, minute, second] = dateMatch;

    if (hour === '24') {
      hour = '0'; // Known issue with toLocaleString() prior to ECMAScript 2021
    }

    return {
      date: {
        year: parseInt(year, 10),
        month: parseInt(month, 10),
        day: parseInt(day, 10),
      },
      hour: parseInt(hour, 10),
      minute: parseInt(minute, 10),
      second: parseInt(second, 10),
      timeZoneId,
    };
  }

  /**
   * Returns a formatted string representation ("YYYY-MM-DD HH:mm:ss") of the
   * provided `DateTime` object for use with the Google Sheets API. The time
   * zone is discarded under the assumption that the spreadsheet's time zone
   * matches the Ad Manager network time zone.
   */
  getDateString(dateTime: ad_manager.DateTime): string {
    const month = String(dateTime.date.month).padStart(2, '0');
    const day = String(dateTime.date.day).padStart(2, '0');
    const hour = String(dateTime.hour).padStart(2, '0');
    const minute = String(dateTime.minute).padStart(2, '0');
    const second = String(dateTime.second).padStart(2, '0');

    return `${dateTime.date.year}-${month}-${day} ${hour}:${minute}:${second}`;
  }

  /**
   * Retrieves line items that are potential candidates for custom delivery
   * curves based on the provided filter and offset.
   *
   * Due to performance limitations that arise inherently from using Apps Script
   * along with SOAP object handling, we need to break the request into smaller
   * batches to improve the user experience. The `offset` parameter is used to
   * page through the results.
   *
   * Notably the `LineItemPage` that is returned is not the same as the one
   * returned from Ad Manager because of additional filtering that cannot be
   * performed in the PQL query. The totalResultSetSize is correct to ensure
   * proper pagination, but the results array may be filtered.
   * @param filter A collection of settings used to filter line items
   * @param offset The number of lines to skip before returning a batch
   * @returns An offset page of line items that match the filter
   */
  getLineItemsByFilter(
    filter: LineItemFilter,
    offset: number,
  ): ad_manager.LineItemPage {
    const whereClause =
      'isArchived = false ' +
      'AND costType = :costType ' +
      'AND deliveryRateType <> :deliveryRateType ' +
      'AND endDateTime >= :endDateTime ' +
      'AND lineItemType = :lineItemType ' +
      'AND startDateTime <= :startDateTime ';

    const statement = new StatementBuilder()
      .where(whereClause)
      .withBindVariable('costType', 'CPM')
      .withBindVariable('deliveryRateType', 'AS_FAST_AS_POSSIBLE')
      .withBindVariable('endDateTime', filter.earliestEndDate.toISOString())
      .withBindVariable('lineItemType', 'STANDARD')
      .withBindVariable('startDateTime', filter.latestStartDate.toISOString())
      .withOffset(offset)
      .withLimit(AdManagerHandler.AD_MANAGER_API_PAGE_LIMIT)
      .toStatement();

    const lineItemService = this.getService('LineItemService');
    const lineItemPage = lineItemService.performOperation(
      'getLineItemsByStatement',
      statement,
    ) as ad_manager.LineItemPage;

    // If no ad unit IDs are provided, then skip the targeting check
    if (filter.adUnitIds.length === 0) {
      return lineItemPage;
    }

    // PQL queries cannot filter on targeting, so handle explicitly here
    const filteredLineItems = lineItemPage.results.filter((lineItem) =>
      this.hasTargetedAdUnitMatch(lineItem, filter.adUnitIds),
    );

    return {
      totalResultSetSize: lineItemPage.totalResultSetSize,
      startIndex: lineItemPage.startIndex,
      results: filteredLineItems,
    };
  }

  /**
   * Retrieves line items from Ad Manager that match the provided IDs.
   * @param ids An array of line item IDs
   * @param offset The number of lines to skip before returning a batch
   */
  getLineItemsWithIds(ids: number[], offset: number): ad_manager.LineItemPage {
    const statement = new StatementBuilder()
      .where('id IN (:ids)')
      .withBindVariable('ids', ids)
      .withOffset(offset)
      .withLimit(AdManagerHandler.AD_MANAGER_API_PAGE_LIMIT)
      .toStatement();

    const lineItemService = this.client.getService('LineItemService');
    return lineItemService.performOperation(
      'getLineItemsByStatement',
      statement,
    ) as ad_manager.LineItemPage;
  }

  /** Returns the time zone of the current network. */
  getTimeZoneId(): string {
    const networkService = this.getService('NetworkService');
    const network = networkService.performOperation(
      'getCurrentNetwork',
    ) as ad_manager.Network;

    return network.timeZone;
  }

  /**
   * Submits a batch of updated line items to the Ad Manager API.
   *
   * Currently this function will allow overbooking and skip inventory checks
   * for all line items in the batch. This should likely be a configurable
   * setting in the future.
   *
   * Ad Manager will reject the entire batch if any line item fails validation.
   * @throws An `AdManagerServerFault` if the operation fails
   */
  uploadLineItems(lineItems: ad_manager.LineItem[]) {
    for (const lineItem of lineItems) {
      lineItem.allowOverbook = true;
      lineItem.skipInventoryCheck = true;
    }

    const lineItemService = this.getService('LineItemService');
    lineItemService.performOperation('updateLineItems', lineItems);
  }

  /**
   * Retrieves a cached or newly created `AdManagerService` by name.
   *
   * This function optimizes service retrieval by maintaining an internal cache.
   * Since each client service can take substantial time to initialize, reusing
   * existing instances where possible significantly improves performance. If a
   * service with the given name isn't cached, a new instance is created and
   * added to the cache for future use.
   * @param serviceName The unique name of the desired Ad Manager service.
   * @returns The `AdManagerService` instance associated with the provided name.
   */
  private getService(serviceName: string): AdManagerService {
    let service = this.serviceCache.get(serviceName);

    if (!service) {
      service = this.client.getService(serviceName);
      this.serviceCache.set(serviceName, service);
    }

    return service;
  }

  /**
   * Returns true if the provided line item targets at least one of the provided
   * ad unit IDs.
   * @param lineItem The line item to check
   * @param adUnitIds An array of ad unit IDs
   */
  private hasTargetedAdUnitMatch(
    lineItem: ad_manager.LineItem,
    adUnitIds: string[],
  ): boolean {
    const targetedAdUnits =
      lineItem.targeting.inventoryTargeting?.targetedAdUnits;

    if (targetedAdUnits) {
      return targetedAdUnits.some(({adUnitId}) => adUnitIds.includes(adUnitId));
    }

    return false;
  }
}
