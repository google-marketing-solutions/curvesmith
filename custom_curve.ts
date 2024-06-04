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
 * @fileoverview This is a library that can be used to generate custom delivery
 * curve data models for upload into Ad Manager.
 */

/** Encapsulates a period of time between two dates. */
export class DateRange {
  readonly start: Date;
  readonly end: Date;

  constructor(start: Date | string, end: Date | string) {
    this.start = new Date(start);
    this.end = new Date(end);

    DateRange.validate(this.start, this.end);
  }

  /**
   * Validates that the provided arguments are valid dates and that the start
   * date precedes the end date.
   * @throws An error if either date is invalid or out of order
   */
  static validate(startDate: Date, endDate: Date): void {
    if (!(startDate instanceof Date) || !(endDate instanceof Date)) {
      throw new TypeError('Input values must be Date objects');
    }

    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      throw new RangeError('Input values must be valid dates');
    }

    if (startDate.getTime() >= endDate.getTime()) {
      throw new RangeError('Start date must strictly precede end date');
    }
  }

  /** Returns true if the provided date range is a subset. */
  contains(dateRange: DateRange): boolean {
    return dateRange.start >= this.start && dateRange.end <= this.end;
  }

  /**
   * Returns true if the provided date range intersects with this one. Ranges
   * are not considered intersecting if they are merely adjacent (i.e. if they
   * share an edge, meaning one range starts immediately after the other ends).
   */
  overlaps(dateRange: DateRange): boolean {
    return this.start < dateRange.end && this.end > dateRange.start;
  }
}

/**
 * An enumeration indicating how event goals should be calculated.
 *
 * - **TOTAL** indicates that goals are relative to the entire flight.
 * - **DAY** indicates that goals are relative to a single 24-hour period.
 */
export enum GoalType {
  TOTAL,
  DAY,
}

/** This is analagous to a line item with an absolute impression goal. */
export class FlightDetails extends DateRange {
  constructor(
    start: Date | string,
    end: Date | string,
    readonly impressionGoal: number,
  ) {
    super(start, end);
  }
}

/**
 * Encapsulates a period for which a user will request a skewed delivery goal.
 */
export class ScheduledEvent extends DateRange {
  readonly DEFAULT_TITLE = 'Untitled';

  constructor(
    start: Date | string,
    end: Date | string,
    readonly goalPercent: number,
    readonly title: string,
  ) {
    super(start, end);
  }

  getTitleForCurve(): string {
    return this.title ? this.title : this.DEFAULT_TITLE;
  }
}

/**
 * A simple data container used to calculate actual goal percentages for upload
 * into Ad Manager once all scheduled events have been processed. This will
 * likely need to expand to accomodate the requirements of additional goal
 * types.
 */
class GoalContext {
  constructor(
    readonly goalType: GoalType,
    readonly impressionGoal: number,
  ) {}

  /**
   * The total number of hours within a flight that are not associated with a
   * scheduled event.
   */
  unscheduledHours = 0;

  /**
   * The total number of impressions within a flight that are not associated
   * with a scheduled event.
   */
  unscheduledImpressions = 0;

  /**
   * The total percentage of hours within a flight that are not associated with
   * a scheduled event.
   */
  unscheduledPercent = 100;
}

/**
 * A segment of a custom curve represents a starting time and a portion of
 * the overall flight goal.
 */
export interface CurveSegment {
  description: string;
  start: Date;
  goalPercent: number;

  /**
   * Updates and returns the goal of this segment relative to the total flight.
   * Calculated goals will be between 0 and 100 (e.g. 0% and 100%).
   */
  calculateGoal(context: GoalContext): number;
}

/** A custom curve segment representing a user-defined period of time. */
class ScheduledEventSegment implements CurveSegment {
  constructor(
    readonly description: string,
    readonly start: Date,
    readonly goalPercent: number,
  ) {}

  calculateGoal() {
    return this.goalPercent;
  }
}

/**
 * A custom curve segment representing a portion of a flight duration that has
 * not been associated with a user-defined event.
 */
class UnscheduledSegment implements CurveSegment {
  goalPercent = 0;

  constructor(
    readonly description: string,
    readonly start: Date,
    private readonly hours: number,
  ) {}

  /**
   * Calculates the goal for this unscheduled time relative to the overall goal.
   */
  calculateGoal(context: GoalContext) {
    // Calculates the proportion of the total unscheduled time within the
    // associated flight that this particular segment represents.
    const unscheduledTimeProportion = this.hours / context.unscheduledHours;

    switch (context.goalType) {
      case GoalType.DAY: {
        const normalizedImpressionCount =
          unscheduledTimeProportion * context.unscheduledImpressions;

        // Normalizes into a percentage of the total impression goal.
        this.goalPercent =
          (normalizedImpressionCount * 100) / context.impressionGoal;
        break;
      }
      case GoalType.TOTAL: {
        // Normalizes into a percentage of the total flight duration.
        this.goalPercent =
          unscheduledTimeProportion * context.unscheduledPercent;
        break;
      }
    }

    return this.goalPercent;
  }
}

const MILLISECONDS_PER_HOUR = 1000 * 60 * 60;

const TOTAL_PERCENT_REQUIRED = 100;

/** Maximum allowable deviation from 100% before a curve error is detected. */
const PERCENT_ERROR_THRESHOLD = 0.001;

/** Defines a template that can be used to generate a custom curve model. */
export class CurveTemplate {
  constructor(
    readonly events: ScheduledEvent[],
    readonly goalType: GoalType = GoalType.TOTAL,
  ) {}

  /**
   * Validate requirements on scheduled events against the provided flight.
   *  - All scheduled events must be within the flight range.
   *  - All scheduled events must be ordered by date.
   *  - No scheduled events may overlap.
   * @throws An error if any of the conditions are not met
   */
  private checkDateRanges(flight: FlightDetails) {
    let dateRange = null;

    for (const event of this.events) {
      if (!flight.contains(event)) {
        throw new Error('Event date range is outside flight range');
      }

      if (dateRange) {
        if (event.start < dateRange.start) {
          throw new Error('Events must be ordered by date');
        }

        if (event.overlaps(dateRange)) {
          throw new Error('Events date ranges must never overlap');
        }
      }

      dateRange = event;
    }
  }

  /**
   * Returns a custom curve for the specified `flight` based upon this template.
   * The curve accounts for scheduled events and fills in the remainder of the
   * flight duration before and after those events as needed.
   *
   * Goal calculation is relative to the total flight duration and outside of
   * explicit skews indicated by scheduled events, the remaining goal is
   * distributed evenly across unscheduled time.
   * @throws An error if a custom delivery curve could not be generated
   */
  generateCurveSegments(flight: FlightDetails): CurveSegment[] {
    this.checkDateRanges(flight);

    let segments: CurveSegment[];
    const goalContext = new GoalContext(this.goalType, flight.impressionGoal);

    switch (this.goalType) {
      case GoalType.DAY:
        segments = this.processEventsByDay(flight, goalContext);
        break;
      case GoalType.TOTAL:
        segments = this.processEventsByTotal(flight, goalContext);
        break;
    }

    // Goals for unscheduled time can only be calculated after partitioning
    // the flight range into curve segments in order to know how what
    // percentage of the flight goal is still available and over how many
    // hours that goal needs to be distributed.
    const totalGoalPercent = segments.reduce(
      (sum, x) => sum + x.calculateGoal(goalContext),
      /** initialValue */ 0,
    );

    // Allow for slight precision adjustment to calculated goals.
    const difference = TOTAL_PERCENT_REQUIRED - totalGoalPercent;

    if (Math.abs(difference) > PERCENT_ERROR_THRESHOLD) {
      throw new Error('Total goal percent must equal 100');
    } else {
      // Add the difference to the last segment goal to ensure the total is
      // exactly 100%; otherwise Ad Manager will reject the curve.
      segments[segments.length - 1].goalPercent += difference;
    }

    return segments;
  }

  /**
   * Process event goals by interpreting values as percentages of what would be
   * a single day's impressions if the total goal were evenly distributed across
   * the flight duration. This allows individual event goals to exceed 100%.
   *
   * Let's say a flight has a 100k impression goal over 10 days: An event with a
   * goal of 300% means it wants 3 times its 'fair share' of daily impressions
   * (which would be 10k in this case), so 30k impressions during its timeframe.
   */
  private processEventsByDay(
    flight: FlightDetails,
    goalContext: GoalContext,
  ): CurveSegment[] {
    const segments: CurveSegment[] = [];

    // Initialize to the total goal and subtract scheduled events
    goalContext.unscheduledImpressions = flight.impressionGoal;

    let currentStart: Date = flight.start;

    const flightHours: number = this.hoursBetween(flight.start, flight.end);

    const evenDailyGoal = (flight.impressionGoal / flightHours) * 24;

    for (const event of this.events) {
      const eventTitle = event.getTitleForCurve();
      const relativeGoal = (event.goalPercent / 100) * evenDailyGoal;
      const normalizedPercent = (relativeGoal / flight.impressionGoal) * 100;
      const hoursBeforeEvent = this.hoursBetween(currentStart, event.start);

      goalContext.unscheduledImpressions -= relativeGoal;

      if (goalContext.unscheduledImpressions <= 0) {
        throw new Error('Goal is too large');
      }

      if (hoursBeforeEvent > 0) {
        goalContext.unscheduledHours += hoursBeforeEvent;

        segments.push(
          new UnscheduledSegment(
            `Pre-Event [${eventTitle}]`,
            currentStart,
            hoursBeforeEvent,
          ),
        );
      }

      segments.push(
        new ScheduledEventSegment(eventTitle, event.start, normalizedPercent),
      );

      goalContext.unscheduledPercent -= normalizedPercent;
      currentStart = event.end;
    }

    const hoursAfterLastEvent = this.hoursBetween(currentStart, flight.end);

    if (hoursAfterLastEvent > 0) {
      // There is time after the last event, but all reserved impressions have
      // been consumed by the events. The final segment of a curve cannot have a
      // zero goal.
      if (goalContext.unscheduledImpressions <= 0) {
        throw new Error('Goal is too large.');
      }

      segments.push(
        new UnscheduledSegment(
          'Post-Events',
          currentStart,
          hoursAfterLastEvent,
        ),
      );

      goalContext.unscheduledHours += hoursAfterLastEvent;
    }

    return segments;
  }

  /**
   * Process event goals by interpreting values as percentages of the entire
   * flight's impression goal. Individual events must be between 0 and 100% and
   * the sum of all event goals must be less than or equal to 100%.
   *
   * Let's say a flight has a 100k impression goal over 10 days: An event with a
   * goal of 30% means it wants 30k impressions delivered during its timeframe.
   */
  private processEventsByTotal(
    flight: FlightDetails,
    goalContext: GoalContext,
  ): CurveSegment[] {
    const segments: CurveSegment[] = [];

    let currentStart: Date = flight.start;

    for (const event of this.events) {
      const eventTitle = event.getTitleForCurve();
      const hoursBeforeEvent = this.hoursBetween(currentStart, event.start);

      if (hoursBeforeEvent > 0) {
        goalContext.unscheduledHours += hoursBeforeEvent;
        segments.push(
          new UnscheduledSegment(
            `Pre-Event [${eventTitle}]`,
            currentStart,
            hoursBeforeEvent,
          ),
        );
      }

      segments.push(
        new ScheduledEventSegment(eventTitle, event.start, event.goalPercent),
      );

      currentStart = event.end;
      goalContext.unscheduledPercent -= event.goalPercent;

      if (goalContext.unscheduledPercent < 0) {
        throw new Error('Total goal percent is greater than 100');
      } else if (goalContext.unscheduledPercent === 0) {
        if (event.end < flight.end) {
          throw new Error('The curve cannot end with a 0 percent goal');
        }
      }
    }

    const hoursAfterLastEvent = this.hoursBetween(currentStart, flight.end);

    if (hoursAfterLastEvent > 0) {
      goalContext.unscheduledHours += hoursAfterLastEvent;
      segments.push(
        new UnscheduledSegment(
          'Post-Events',
          currentStart,
          hoursAfterLastEvent,
        ),
      );
    }

    return segments;
  }

  /** Returns the number of hours between two dates. */
  private hoursBetween(start: Date, end: Date): number {
    return (end.valueOf() - start.valueOf()) / MILLISECONDS_PER_HOUR;
  }
}
