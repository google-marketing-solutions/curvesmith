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
 * @fileoverview Typing for the Ad Manager API.
 *
 * To facilitate the use of the GAM Apps Script library, we need to declare
 * interfaces that mirror the Ad Manager API object model for any objects that
 * we want to access. We only need to explicitly define properties that we will
 * access as part of this solution, so many of the following interfaces are
 * declared as only partial definitions.
 */

/** Represents an Ad Manager API `AdUnit` object. */
export declare interface AdUnit {
  id: string;
  parentId: string;
  hasChildren: boolean;
  explicitlyTargeted: boolean;
}

/**
 * Represents an `AdUnitPage` object returned from the Ad Manager API when
 * requesting a collection of ad units through `getAdUnitsByStatement`.
 */
export declare interface AdUnitPage {
  totalResultSetSize: number;
  startIndex: number;
  results: AdUnit[];
}

/** Represents an Ad Manager API `AdUnitTargeting` object. */
export declare interface AdUnitTargeting {
  adUnitId: string;
}

/** Represents an Ad Manager API `CustomPacingCurve` object. */
export declare interface CustomPacingCurve {
  customPacingGoalUnit: string;
  customPacingGoals: CustomPacingGoal[];
}

/** Represents an Ad Manager API `CustomPacingGoal` object. */
export declare interface CustomPacingGoal {
  startDateTime?: DateTime;
  useLineItemStartDateTime: boolean;
  amount: number;
}

/** Represents an Ad Manager API `Date` object. */
export declare interface Date {
  year: number;
  month: number;
  day: number;
}

/** Represents an Ad Manager API `DateTime` object. */
export declare interface DateTime {
  date: Date;
  hour: number;
  minute: number;
  second: number;
  timeZoneId: string;
}

/** Represents an Ad Manager API `Goal` object. */
export declare interface Goal {
  /**
   * This solution is only relevant to STANDARD line items, so the `units`
   * property always refers to an explicit impression number.
   */
  units: number;
}

/** Represents an Ad Manager API `InventoryTargeting` object. */
export declare interface InventoryTargeting {
  targetedAdUnits?: AdUnitTargeting[];
}

/** Represents an Ad Manager API `LineItem` object. */
export declare interface LineItem {
  id: number;
  name: string;
  startDateTime: DateTime;
  endDateTime: DateTime;
  autoExtensionDays: number;
  deliveryForecastSource?: string;
  customPacingCurve?: CustomPacingCurve;
  allowOverbook?: boolean;
  skipInventoryCheck?: boolean;
  primaryGoal: Goal;
  targeting: Targeting;
}

/**
 * Represents a `LineItemPage` object returned from the Ad Manager API when
 * requesting a collection of lines through `getLineItemsByStatement`.
 */
export declare interface LineItemPage {
  totalResultSetSize: number;
  startIndex: number;
  results: LineItem[];
}

/** Represents an Ad Manager API `Network` object. */
export declare interface Network {
  timeZone: string;
}

/**
 * Represents an Ad Manager API `Targeting` object. Currently we are only
 * interested in `inventoryTargeting` so we can filter by ad unit.
 */
export declare interface Targeting {
  inventoryTargeting: InventoryTargeting;
}

/** Interfaces specific to the PublisherQueryLanguageService. */

export declare interface ColumnType {
  labelName: string;
}

export declare interface ResultSet {
  columnTypes: ColumnType[];
  rows: Row[];
}

export declare interface Row {
  values: Value[];
}

export declare interface Value {
  value: boolean | Date | DateTime | number | string | Targeting;
}
