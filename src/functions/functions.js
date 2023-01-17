/**
 * @customfunction GetStreamValue
 * @param {string} streamNumber
 * @param {string} operatingCase
 * @param {string} propertyType
 * @param {string[][]} searchRange
 * @returns {string} of PropertyValue.
 */
export function getStreamValue(streamNumber, operatingCase, propertyType, searchRange) {
  const streamNumberIndex = findHeaderIndex("StreamNumber", searchRange);
  const operatingCaseIndex = findHeaderIndex("OperatingCase", searchRange);
  const propertyTypeIndex = findHeaderIndex("PropertyType", searchRange);
  const propertyValueIndex = findHeaderIndex("PropertyValue", searchRange);

  const matchingRow = searchRange.find(
    (element) =>
      element[streamNumberIndex] === streamNumber &&
      element[operatingCaseIndex] === operatingCase &&
      element[propertyTypeIndex] === propertyType
  );

  if (!matchingRow) {
    let error = new CustomFunctions.Error(
      CustomFunctions.ErrorCode.notAvailable,
      `Could not find PropertyValue for StreamNumber: ${streamNumber}, OperatingCase: ${operatingCase}, PropertyType: ${propertyType}`
    );
    throw error;
  }

  return matchingRow[propertyValueIndex];
}

const findHeaderIndex = (headerName, searchRange) => {
  const index = searchRange[0].findIndex((element) => element === headerName);
  if (index === -1) {
    let error = new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `${headerName} not found in header row (first row of searchRange)`
    );
    throw error;
  }
  return index;
};
