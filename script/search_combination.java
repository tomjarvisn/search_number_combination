/**
 * @customfunction
 */
function SEARCH_COMBINATION(cellRange, targetCell, marginOfError, positiveAndNegativeValues) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  try {
    // Get cell values only once
    var data = sheet.getRange(cellRange).getValues().flat();

    // Filter values based on the positiveAndNegativeValues parameter
    data = data.filter(function(value) {
      if (positiveAndNegativeValues) {
        return value !== null;
      } else {
        return value !== null && value !== 0;
      }
    });

    // Filter values based on the relationship with the target
    if (!positiveAndNegativeValues) {
      var targetValue = sheet.getRange(targetCell).getValue();
      data = data.filter(function(value) {
        return (targetValue > 0 && value <= targetValue) || (targetValue < 0 && value >= targetValue);
      });
    }

    // Get the value of the target cell
    var targetCellValue = sheet.getRange(targetCell).getValue();

    // Convert values to numbers and sort the range (optional)
    var numericData = data.map(function(cell) {
      return parseFloat(cell);
    }).sort(function(a, b) {
      return a - b;
    });

    // Call the recursiveSearch function with the data, target value, margin of error, and positiveAndNegativeValues parameter
    var foundCombinations = recursiveSearch(numericData, targetCellValue, 0, 0, [], marginOfError, positiveAndNegativeValues);

    // Return the information in transposed form
    if (foundCombinations.length > 0) {
      return transposeMatrix(foundCombinations);
    } else {
      return "No combinations found";
    }
  } catch (error) {
    console.error(error);
    return "Error: " + error.message;
  }
}

/**
 * Function to recursively search for combinations.
 */
function recursiveSearch(data, target, accumulated, index, currentCombination, marginOfError, positiveAndNegativeValues) {
  currentCombination = currentCombination || [];
  marginOfError = marginOfError || 0;
  positiveAndNegativeValues = positiveAndNegativeValues || false;

  function isValidValue(value) {
    if (positiveAndNegativeValues) {
      return value !== null;
    } else {
      return value !== null && value !== 0;
    }
  }

  if (Math.abs(accumulated - target) <= marginOfError && currentCombination.length > 0) {
    return [currentCombination];
  }

  if (index >= data.length) {
    return [];
  }

  // Filter values based on positiveAndNegativeValues parameter and relationship with the target
  var currentValue = data[index];
  if (positiveAndNegativeValues) {
    var include = isValidValue(currentValue) ?
      recursiveSearch(data, target, accumulated + currentValue, index + 1, [...currentCombination, currentValue], marginOfError, positiveAndNegativeValues) :
      [];
  } else {
    var condition = (target > 0 && currentValue <= target) || (target < 0 && currentValue >= target);
    var include = condition && isValidValue(currentValue) ?
      recursiveSearch(data, target, accumulated + currentValue, index + 1, [...currentCombination, currentValue], marginOfError, positiveAndNegativeValues) :
      [];
  }

  // Do not include the current element in the combination
  var exclude = recursiveSearch(data, target, accumulated, index + 1, currentCombination, marginOfError, positiveAndNegativeValues);

  // Combine the results
  return include.concat(exclude);
}

/**
 * Function to transpose a matrix.
 */
function transposeMatrix(matrix) {
  return matrix[0].map(function(_, i) {
    return matrix.map(function(row) {
      return row[i];
    });
  });
}
