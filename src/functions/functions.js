/**
 * @customfunction GetStreamValue
 * @param {number} streamNumber
 * @param {string} operatingCase
 * @param {string} propertyType
 * @returns {string} of PropertyValue.
 */

export async function getStreamValue(streamNumber, operatingCase, propertyType) {
  const searchRange = await getRange();

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

const getRange = async () => {
  try {
    const context = new Excel.RequestContext();
    const worksheet = context.workbook.worksheets.getItem("Test");
    const range = worksheet.getRange("PropTypeValue");
    const searchRange = range.load("values");
    await context.sync();
    return searchRange.values;
  } catch (searchError) {
    let error = new CustomFunctions.Error(
      CustomFunctions.ErrorCode.invalidValue,
      `Search range not found with following error: ${searchError.message}`
    );
    throw error;
  }
};

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
