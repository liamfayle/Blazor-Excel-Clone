function getInputValueByDataAttributes(dataX, dataY) {
    const element = document.querySelector(`input[data-x='${dataX}'][data-y='${dataY}']`);
    return element ? element.value : null;
}

function setInputValueByDataAttributes(dataX, dataY, value) {
    const element = document.querySelector(`input[data-x='${dataX}'][data-y='${dataY}']`);
    element ? element.value : null;
    if (element.value != null) {
        element.value = value;
        element.blur();
    }
}