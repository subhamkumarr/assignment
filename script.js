document.addEventListener("DOMContentLoaded", function () {
    var submitButton = document.getElementById("submitButton");
    var downloadButton = document.getElementById("downloadButton"); 
    var displayArea = document.getElementById("displayArea");
    var serialNumberMap = new Map();

    loadMachineData();

    submitButton.addEventListener("click", function () {
        var machineNumber = document.getElementById("machineNumber").value;

        if (serialNumberMap.has(machineNumber)) {
            var serialNumber = serialNumberMap.get(machineNumber);
        } else {
            var serialNumber = '1';

            serialNumberMap.set(machineNumber, serialNumber);
        }

        var newRow = document.createElement("tr");
        newRow.innerHTML = `
            <td>${serialNumber}</td>
            <td>${getCurrentDate()}</td>
            <td>${getCurrentTime()}</td>
            <td>${machineNumber}</td>
        `;

        var table = displayArea.querySelector("table");
        table.appendChild(newRow);

        document.getElementById("machineNumber").value = "";

        saveMachineData();

        displayArea.style.display = "block";
    });

    downloadButton.addEventListener("click", function () {
        var table = displayArea.querySelector("table");
        exportTableToExcel(table, "MachineData"); 
    });

    function loadMachineData() {
        var data = localStorage.getItem('serialNumberMap');
        if (data) {
            serialNumberMap = new Map(JSON.parse(data));
        }
    }

    function saveMachineData() {
        localStorage.setItem('serialNumberMap', JSON.stringify(Array.from(serialNumberMap.entries())));
    }

    function getCurrentDate() {
        var currentDate = new Date();
        return currentDate.toDateString();
    }

    function getCurrentTime() {
        var currentDate = new Date();
        return currentDate.toLocaleTimeString();
    }

    function exportTableToExcel(table, filename) {
        var downloadLink;
        var dataType = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
        var tableHTML = table.outerHTML.replace(/ /g, '%20');

        downloadLink = document.createElement('a');
        document.body.appendChild(downloadLink);

        if (navigator.msSaveOrOpenBlob) {
            var blob = new Blob(['\ufeff', tableHTML], {
                type: dataType
            });
            navigator.msSaveOrOpenBlob(blob, filename + '.xls');
        } else {
            downloadLink.href = 'data:' + dataType + ', ' + tableHTML;

            downloadLink.download = filename + '.xls';

            downloadLink.click();
        }
    }
});
