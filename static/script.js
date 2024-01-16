function validateForm() {
    console.log("validateForm has been called!!!!!!!!!!!!!!!!")
    var fileInput = document.getElementById('input_file');
    var sheetSelection = document.querySelector('input[name="fav_language"]:checked').value;

    if (!fileInput.files.length) {
        alert("Please select an Excel file.");
        return false;
    }

    var fileName = fileInput.files[0].name;

    if (!isExcelFile(fileName)) {
        alert("Invalid file type. Please upload an Excel file.");
        return false;
    }

    var isValid = true;

    if (!((sheetSelection === "PL1" && containsValidKeyword(fileName, "KCP1")) ||
        (sheetSelection === "XXXX" && containsValidKeyword(fileName, "XXXX")))) {
        alert("Invalid file name. Please select the correct radio button for the file.");
        isValid = false;
    }

    var reader = new FileReader();
    reader.onload = function (e) {
        var data = new Uint8Array(e.target.result);
        var workbook = XLSX.read(data, { type: 'array' });

        if (!workbook.SheetNames.includes("Samples") ||
            !workbook.SheetNames.includes("High Controls")) {
            alert("The Excel file is missing required sheets. Please check the contents.");
            isValid = false;
        }

        if (isValid) {
            // Submit the form if all checks pass
            document.forms[0].submit();
        }
    };

    reader.readAsArrayBuffer(fileInput.files[0]);

    return false;
}

function isExcelFile(filename) {
    return filename.endsWith('.xlsx');
}

function containsValidKeyword(filename, keyword) {
    return filename.includes(keyword);
}

document.addEventListener('DOMContentLoaded', function () {
    var radioButtons = document.querySelectorAll('input[name="fav_language"]');
    var hiddenInput = document.getElementById('selected_radio_id');

    // Set the initial value
    hiddenInput.value = document.querySelector('input[name="fav_language"]:checked').id;

    // Add an event listener to update the value when the radio button changes
    radioButtons.forEach(function (radioButton) {
        radioButton.addEventListener('change', function () {
            hiddenInput.value = this.id;
        });
    });
});
