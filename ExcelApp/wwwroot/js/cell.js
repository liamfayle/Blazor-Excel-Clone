
document.addEventListener('DOMContentLoaded', () => {
    document.querySelectorAll('.cell input').forEach(element => {

        element.addEventListener('keydown', function (e) {
            if (e.key === "Enter") {
                // Prevent default action if necessary, such as form submission
                e.preventDefault();

                // Process mathematical operations
                const value = this.value;
                if (value.startsWith('=')) {
                    try {
                        const stringToProcess = getStringToProcess(value.substring(1));
                        const result = eval(stringToProcess);

                        if (result)
                            this.value = result;
                        else {
                            this.value = ""
                        }
                    } catch (error) {
                        console.error('Error processing the expression:', error);
                        this.value = 'Error';
                    }
                }

                this.blur();
            }
        });
    });
});

function extractIndividualLetterNumberPairs(string) {
    // Regular expression to match individual letter-number pairs
    // Excludes those that are part of a range (e.g., 'A1:B2')
    const regex = /(?<![:A-Za-z]\d+)(?<![:A-Za-z])[A-Za-z]\d+\b(?![:A-Za-z]\d+)/g;

    // Find all matches in the string
    const matches = string.match(regex);

    // If there are matches, return them; otherwise, return an empty array
    return matches || [];
}

function extractRangesInFunctions(string) {
    // Regular expression to match ranges surrounded by SUM() or AVERAGE() including the function
    const regex = /(sum|average)\(([A-Za-z]\d+:[A-Za-z]\d+)\)/gi;

    // Find all matches in the string
    const matches = [];
    let match;
    while ((match = regex.exec(string)) !== null) {
        // Add the entire matched string (function with range)
        matches.push(match[0]); // match[0] contains the whole matched string like 'sum(b66:b99)'
    }

    return matches;
}

//function to replace letters with values
function getStringToProcess(string) {
    string = string.toLowerCase();
    const individual = extractIndividualLetterNumberPairs(string);
    const functions = extractRangesInFunctions(string);

    console.log(individual, functions)

    //process functions
    for (const func of functions) {
        var range = func.replace('sum(', '').replace('average(', '').replace(')', '').split(":");
        const { x: x1, y: y1 } = convertExcelToXY(range[0]);
        const { x: x2, y: y2 } = convertExcelToXY(range[1]);

        var value = loopThroughRange(x1, y1, x2, y2, string.substring(0, 3));

        string = string.replace(func, value + "");
    }

    for (const ind of individual) {
        const { x, y } = convertExcelToXY(ind)
        var val = getValueXY(x, y);
        if (val != null) {
            string = string.replace(ind, val)
        } else {
            string = string.replace(ind, "")
        }
    }

    return string;
}


function splitStringIntoLetterAndNumberPart(inputString) {
    // Regular expression to separate letters and numbers
    const regex = /^([A-Za-z]+)(\d+)$/;
    const match = inputString.match(regex);

    if (match) {
        return { letter: match[1], number: match[2] };
    } else {
        return null; // Or throw an error, or return inputString as fallback
    }
}

function convertExcelToXY(excel) {
    const { letter, number } = splitStringIntoLetterAndNumberPart(excel);

    // Convert the letter part to lowercase and then get the first character's code
    // minus the character code for 'a' to get the index (0-based)
    const letterIndex = letter.toLowerCase().charCodeAt(0) - 'a'.charCodeAt(0);

    return {x: letterIndex, y: number-1}
}


function loopThroughRange(x1, y1, x2, y2, func) {
    // Ensure the starting point is the smaller one
    const startX = Math.min(x1, x2);
    const endX = Math.max(x1, x2);
    const startY = Math.min(y1, y2);
    const endY = Math.max(y1, y2);

    var count = 0;
    var total = 0;

    // Loop from start to end indices
    for (let x = startX; x <= endX; x++) {
        for (let y = startY; y <= endY; y++) {
            var val = getValueXY(x, y);
            if (val != null) {
                total += parseFloat(val);
                count++;
            }
        }
    }

    if (func == "sum") {
        return total;
    }
    return total / count;
}


function getValueXY(x, y) {
    // Use the attributes in the query selector to find the input element
    var inputElement = document.querySelector(`input[data-x="${x}"][data-y="${y}"]`);

    // Get the value of the input element, if it exists
    var value = inputElement ? inputElement.value : null;

    if (isNumeric(value)) {
        return value;
    } 
    return null
}


function isNumeric(str) {
    if (typeof str != "string") return false // we only process strings!  
    return !isNaN(str) && // use type coercion to parse the _entirety_ of the string (`parseFloat` alone does not do this)...
        !isNaN(parseFloat(str)) // ...and ensure strings of whitespace fail
}