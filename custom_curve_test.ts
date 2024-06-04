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
  CurveTemplate,
  DateRange,
  FlightDetails,
  GoalType,
  ScheduledEvent,
} from './custom_curve';

describe('DateRange', () => {
  describe('.ctor', () => {
    it('throws an error if arguments are invalid text', () => {
      expect(() => {
        new DateRange('invalid', 'invalid');
      }).toThrow(new RangeError('Input values must be valid dates'));
    });
  });

  describe('validate', () => {
    it('throws an error if arguments represent infinite values', () => {
      expect(() => {
        DateRange.validate(new Date(Infinity), new Date(Infinity));
      }).toThrow(new RangeError('Input values must be valid dates'));
    });

    it('throws an error if start date does not precede end date', () => {
      expect(() => {
        DateRange.validate(new Date('2024-01-31'), new Date('2024-01-30'));
      }).toThrow(new RangeError('Start date must strictly precede end date'));
    });
  });

  describe('.contains(x)', () => {
    it('returns true if x is a subset of this', () => {
      const firstRange = new DateRange('2024-01-01', '2024-06-01');
      const secondRange = new DateRange('2024-02-01', '2024-05-01');

      expect(firstRange.contains(secondRange)).toBe(true);
    });

    it('returns false if x is only a partial subset of this', () => {
      const firstRange = new DateRange('2024-01-01', '2024-05-01');
      const secondRange = new DateRange('2024-02-01', '2024-06-01');

      expect(firstRange.contains(secondRange)).toBe(false);
    });

    it('returns false if x is not a subset of this', () => {
      const firstRange = new DateRange('2024-02-01', '2024-05-01');
      const secondRange = new DateRange('2024-01-01', '2024-06-01');

      expect(firstRange.contains(secondRange)).toBe(false);
    });
  });

  describe('.overlaps(x)', () => {
    it('returns true if the range of x intersects', () => {
      const firstRange = new DateRange('2024-01-01', '2024-04-01');
      const secondRange = new DateRange('2024-02-01', '2024-05-01');

      expect(firstRange.overlaps(secondRange)).toBe(true);
    });

    it('returns true regardless of order', () => {
      const firstRange = new DateRange('2024-01-01', '2024-04-01');
      const secondRange = new DateRange('2024-02-01', '2024-05-01');

      expect(firstRange.overlaps(secondRange)).toBe(
        secondRange.overlaps(firstRange),
      );
    });

    it('returns false regardless of order', () => {
      const firstRange = new DateRange('2024-01-01', '2024-02-01');
      const secondRange = new DateRange('2024-03-01', '2024-04-01');

      expect(firstRange.overlaps(secondRange)).toBe(
        secondRange.overlaps(firstRange),
      );
    });

    it('returns false if the range of x does not intersect', () => {
      const firstRange = new DateRange('2024-01-01', '2024-02-01');
      const secondRange = new DateRange('2024-03-01', '2024-04-01');

      expect(firstRange.overlaps(secondRange)).toBe(false);
    });

    it('returns false if x is adjacent to this', () => {
      const firstRange = new DateRange('2024-01-01', '2024-02-01');
      const secondRange = new DateRange('2024-02-01', '2024-03-01');

      expect(firstRange.overlaps(secondRange)).toBe(false);
    });
  });
});

describe('CurveTemplate', () => {
  describe('.generateCurveSegments()', () => {
    it('throws an error if any events are outside the flight range', () => {
      const flight = new FlightDetails('1/2/2024', '1/2/2024 12:00:00', 10000);
      const events = [
        new ScheduledEvent('1/1/2024 06:00:00', '1/3/2024 09:00:00', 10, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Event date range is outside flight range');
    });

    it('throws an error if any events overlap', () => {
      const flight = new FlightDetails('1/1/2024', '1/5/2024', 10000);
      const events = [
        new ScheduledEvent('1/1/2024', '1/3/2024 10:00:00', 10, 'A'),
        new ScheduledEvent('1/2/2024', '1/3/2024 10:00:00', 10, 'B'),
      ];
      const curveTemplate = new CurveTemplate(events);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Events date ranges must never overlap');
    });

    it('throws an error if event dates are unordered', () => {
      const flight = new FlightDetails('1/1/2024', '3/1/2024', 10000);
      const events = [
        new ScheduledEvent('1/5/2024 06:00:00', '1/5/2024 09:00:00', 10, 'A'),
        new ScheduledEvent('1/5/2024 03:00:00', '1/5/2024 05:00:00', 10, 'B'),
      ];
      const curveTemplate = new CurveTemplate(events);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Events must be ordered by date');
    });
  });

  describe('[GoalType.DAY] .generateCurveSegments()', () => {
    it('returns a curve for a single multi-day event after start', () => {
      const flight = new FlightDetails('3/27/2024', '4/1/2024 23:59:00', 10000);
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 80, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [A]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 12.561679905035467,
        }),
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 20:00:00'),
          goalPercent: 13.334876721842809,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/28/2024 02:00:00'),
          goalPercent: 74.10344337312172,
        }),
      ]);
    });

    it('returns a curve for a single multi-day event from start', () => {
      const flight = new FlightDetails(
        '3/27/2024 20:00:00',
        '4/1/2024 23:59:00',
        10000,
      );
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 80, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 20:00:00'),
          goalPercent: 15.485952412958731,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/28/2024 02:00:00'),
          goalPercent: 84.51404758704128,
        }),
      ]);
    });

    it('returns a curve for two multi-day events after start', () => {
      const flight = new FlightDetails('3/27/2024', '4/1/2024 23:59:00', 10000);
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 80, 'A'),
        new ScheduledEvent('3/29/2024 20:00:00', '3/30/2024 02:00:00', 80, 'B'),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [A]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 11.112046453791796,
        }),
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 20:00:00'),
          goalPercent: 13.334876721842809,
        }),
        jasmine.objectContaining({
          description: 'Pre-Event [B]',
          start: new Date('3/28/2024 02:00:00'),
          goalPercent: 23.335297552962775,
        }),
        jasmine.objectContaining({
          description: 'B',
          start: new Date('3/29/2024 20:00:00'),
          goalPercent: 13.334876721842809,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/30/2024 02:00:00'),
          goalPercent: 38.88290254955981,
        }),
      ]);
    });

    it('returns a curve for two multi-day events from start', () => {
      const flight = new FlightDetails(
        '3/27/2024 20:00:00',
        '4/1/2024 23:59:00',
        10000,
      );
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 80, 'A'),
        new ScheduledEvent('3/29/2024 20:00:00', '3/30/2024 02:00:00', 80, 'B'),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 20:00:00'),
          goalPercent: 15.485952412958731,
        }),
        jasmine.objectContaining({
          description: 'Pre-Event [B]',
          start: new Date('3/28/2024 02:00:00'),
          goalPercent: 25.8893882778223,
        }),
        jasmine.objectContaining({
          description: 'B',
          start: new Date('3/29/2024 20:00:00'),
          goalPercent: 15.485952412958731,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/30/2024 02:00:00'),
          goalPercent: 43.138706896260246,
        }),
      ]);
    });

    it('returns a curve despite an event goal exceeding 100%', () => {
      const flight = new FlightDetails('3/27/2024', '4/1/2024 23:59:00', 10000);
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 200, ''),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [Untitled]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 9.662443511833743,
        }),
        jasmine.objectContaining({
          description: 'Untitled',
          start: new Date('3/27/2024 20:00:00'),
          goalPercent: 33.337191804607016,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/28/2024 02:00:00'),
          goalPercent: 57.00036468355923,
        }),
      ]);
    });

    it('throws an error if event forces total percent less than 100', () => {
      const flight = new FlightDetails('3/27/2024', '3/27/2024 12:00:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024', '3/27/2024 12:00:00', 10, ''),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Total goal percent must equal 100');
    });

    it('throws an error if the event goal is too large once normalized', () => {
      const flight = new FlightDetails('3/27/2024', '4/1/2024 23:59:00', 10000);
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024', 900, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Goal is too large');
    });

    it('throws an error if final curve segment has a zero goal percent', () => {
      const flight = new FlightDetails(
        '3/27/2024 12:00:00',
        '4/1/2024 23:59:00',
        1000000,
      );
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024', 599.931, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events, GoalType.DAY);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Goal is too large');
    });
  });

  describe('[GoalType.TOTAL] .generateCurveSegments()', () => {
    it('returns curve when calculated goal is initially 100.001%', () => {
      const flight = new FlightDetails('4/3/2024', '7/1/2024', 1000000);
      const events = [
        new ScheduledEvent('4/4/2024', '4/8/2024', 10, 'Week 1'),
        new ScheduledEvent('4/11/2024', '4/15/2024', 10, 'Week 2'),
        new ScheduledEvent('4/18/2024', '4/22/2024', 10, 'Week 3'),
        new ScheduledEvent('4/25/2024', '4/29/2024', 10, 'Week 4'),
      ];
      const curveTemplate = new CurveTemplate(events);
      const segments = curveTemplate.generateCurveSegments(flight);

      const goalTotal = segments.reduce((sum, x) => sum + x.goalPercent, 0);

      expect(goalTotal).toEqual(100);
    });

    it('returns curve when calculated goal is initially 99.999%', () => {
      const flight = new FlightDetails('4/3/2024', '7/1/2024', 1000000);
      const events = [
        new ScheduledEvent('4/4/2024', '4/8/2024', 10, 'Week 1'),
        new ScheduledEvent('4/11/2024', '4/15/2024', 10, 'Week 2'),
        new ScheduledEvent('4/18/2024', '4/22/2024', 10, 'Week 3'),
        new ScheduledEvent('4/25/2024', '4/29/2024', 10, 'Week 4'),
        new ScheduledEvent('5/2/2024', '5/6/2024', 10, 'Week 5'),
        new ScheduledEvent('5/9/2024', '5/13/2024', 10, 'Week 6'),
        new ScheduledEvent('5/16/2024', '5/20/2024', 10, 'Week 7'),
      ];
      const curveTemplate = new CurveTemplate(events);
      const segments = curveTemplate.generateCurveSegments(flight);

      const goalTotal = segments.reduce((sum, x) => sum + x.goalPercent, 0);

      expect(goalTotal).toEqual(100);
    });

    it('returns a curve covering the entire flight', () => {
      const flight = new FlightDetails('3/27/2024', '3/27/2024 12:00:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024', '3/27/2024 12:00:00', 100, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024'),
          goalPercent: 100,
        }),
      ]);
    });

    it('returns a curve for a partial-day flight', () => {
      const flight = new FlightDetails('3/27/2024', '3/27/2024 12:00:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024 06:00:00', '3/27/2024 09:00:00', 20, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [A]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 53.33333333333333,
        }),
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 06:00:00'),
          goalPercent: 20,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/27/2024 09:00:00'),
          goalPercent: 26.666666666666664,
        }),
      ]);
    });

    it('returns a curve for a multi-day flight', () => {
      const flight = new FlightDetails('3/27/2024', '3/29/2024 23:59:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024 06:00:00', '3/27/2024 09:00:00', 20, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [A]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 6.958202464363373,
        }),
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 06:00:00'),
          goalPercent: 20,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/27/2024 09:00:00'),
          goalPercent: 73.04179753563662,
        }),
      ]);
    });

    it('returns a curve for multiple events', () => {
      const flight = new FlightDetails('3/27/2024', '3/29/2024 23:59:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024 06:00:00', '3/27/2024 09:00:00', 20, 'A'),
        new ScheduledEvent('3/27/2024 12:00:00', '3/27/2024 16:00:00', 30, 'B'),
      ];
      const curveTemplate = new CurveTemplate(events);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [A]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 4.616568350859194,
        }),
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 06:00:00'),
          goalPercent: 20,
        }),
        jasmine.objectContaining({
          description: 'Pre-Event [B]',
          start: new Date('3/27/2024 09:00:00'),
          goalPercent: 2.308284175429597,
        }),
        jasmine.objectContaining({
          description: 'B',
          start: new Date('3/27/2024 12:00:00'),
          goalPercent: 30,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/27/2024 16:00:00'),
          goalPercent: 43.075147473711205,
        }),
      ]);
    });

    it('returns a curve for an event that spans multiple days', () => {
      const flight = new FlightDetails('3/27/2024', '3/29/2024 23:59:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 30, 'A'),
      ];
      const curveTemplate = new CurveTemplate(events);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [A]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 21.217479161404395,
        }),
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 20:00:00'),
          goalPercent: 30,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/28/2024 02:00:00'),
          goalPercent: 48.7825208385956,
        }),
      ]);
    });

    it('returns a curve for two events spanning multiple days', () => {
      const flight = new FlightDetails('3/27/2024', '3/29/2024 23:59:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024 20:00:00', '3/28/2024 02:00:00', 30, 'A'),
        new ScheduledEvent('3/28/2024 20:00:00', '3/29/2024 02:00:00', 30, 'B'),
      ];
      const curveTemplate = new CurveTemplate(events);

      const segments = curveTemplate.generateCurveSegments(flight);

      expect(segments).toEqual([
        jasmine.objectContaining({
          description: 'Pre-Event [A]',
          start: new Date('3/27/2024 00:00:00'),
          goalPercent: 13.33703806612948,
        }),
        jasmine.objectContaining({
          description: 'A',
          start: new Date('3/27/2024 20:00:00'),
          goalPercent: 30,
        }),
        jasmine.objectContaining({
          description: 'Pre-Event [B]',
          start: new Date('3/28/2024 02:00:00'),
          goalPercent: 12.003334259516532,
        }),
        jasmine.objectContaining({
          description: 'B',
          start: new Date('3/28/2024 20:00:00'),
          goalPercent: 30,
        }),
        jasmine.objectContaining({
          description: 'Post-Events',
          start: new Date('3/29/2024 02:00:00'),
          goalPercent: 14.659627674353988,
        }),
      ]);
    });

    it('throws an error if event forces total percent less than 100', () => {
      const flight = new FlightDetails('3/27/2024', '3/27/2024 12:00:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024', '3/27/2024 12:00:00', 10, ''),
      ];
      const curveTemplate = new CurveTemplate(events);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Total goal percent must equal 100');
    });

    it('throws an error if total goal percentage is greater than 100', () => {
      const flight = new FlightDetails('3/27/2024', '3/27/2024 12:00:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024 06:00:00', '3/27/2024 09:00:00', 80, ''),
        new ScheduledEvent('3/27/2024 09:00:00', '3/27/2024 12:00:00', 90, ''),
      ];
      const curveTemplate = new CurveTemplate(events);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('Total goal percent is greater than 100');
    });

    it('throws an error if final curve segment has a zero goal percent', () => {
      const flight = new FlightDetails('3/27/2024', '3/27/2024 12:00:00', 100);
      const events = [
        new ScheduledEvent('3/27/2024 06:00:00', '3/27/2024 09:00:00', 100, ''),
      ];
      const curveTemplate = new CurveTemplate(events);

      expect(() => {
        curveTemplate.generateCurveSegments(flight);
      }).toThrowError('The curve cannot end with a 0 percent goal');
    });
  });
});
