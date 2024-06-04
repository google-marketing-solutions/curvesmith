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

import {formatErrorMessage} from './common';

describe('common', () => {
  describe('formatErrorMessage', () => {
    it('formats Ad Manager error with trigger', () => {
      const originalMessage =
        "AdManagerServerFault: [AuthenticationError.NETWORK_NOT_FOUND @ ; trigger:'ABC']";

      const errorMessage = formatErrorMessage(new Error(originalMessage));

      expect(errorMessage).toEqual('Ad Manager Error: NETWORK_NOT_FOUND (ABC)');
    });

    it('formats Ad Manager error without trigger', () => {
      const originalMessage =
        'AdManagerServerFault: [AuthenticationError.NO_NETWORKS_TO_ACCESS @ ]';

      const errorMessage = formatErrorMessage(new Error(originalMessage));

      expect(errorMessage).toEqual('Ad Manager Error: NO_NETWORKS_TO_ACCESS');
    });

    it('pass-through non-Ad Manager error without formatting', () => {
      const originalMessage = 'Error: Something else went wrong';

      const errorMessage = formatErrorMessage(new Error(originalMessage));

      expect(errorMessage).toEqual(originalMessage);
    });
  });
});
