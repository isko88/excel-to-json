let selectedFile;
console.log(window.XLSX);
document.getElementById('input').addEventListener("change", (event) => {
    selectedFile = event.target.files[0];
})

let data = [{
    "name": "jayanth",
    "data": "scd",
    "abc": "sdef"
}]

/**
 * 
 * [
    {
        "Order number": 1052,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Dingo Keina",
        "Package size": "Large letter",
        "Tracking number": "1108361BAF2          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1051,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "D Hourst",
        "Package size": "Large letter",
        "Tracking number": "1108361BAE3          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1050,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Steve Newnham",
        "Package size": "Large letter",
        "Tracking number": "1108361BAD4          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1049,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Tom Mortimore",
        "Package size": "Large letter",
        "Tracking number": "1108361BAC8          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1048,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Tom Mortimore",
        "Package size": "Large letter",
        "Tracking number": "1108361BABA          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1047,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Tom Mortimore",
        "Package size": "Large letter",
        "Tracking number": "1108361BAAC          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1046,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Tom Mortimore",
        "Package size": "Large letter",
        "Tracking number": "1108361BA9E          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1036,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Mark Howarth",
        "Package size": "Large letter",
        "Tracking number": "1108361B9F3          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1035,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Liam Watson",
        "Package size": "Large letter",
        "Tracking number": "1108361B9E4          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1034,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "B Langdon",
        "Package size": "Large letter",
        "Tracking number": "1108361B9D5          ",
        "Price paid": 2.05
    },
    {
        "Order number": 1049,
        "Channel": "Manual Order Entry ",
        "Despatch date": 44767.13843318939,
        "Customer": "Tom Mortimore",
        "Package size": "Large letter",
        "Tracking number": "1108361BAC9          ",
        "Price paid": 2.05
    }
]
 */
document.getElementById('button').addEventListener("click", () => {
    XLSX.utils.json_to_sheet(data, 'out.xlsx');
    if (selectedFile) {
        let fileReader = new FileReader();
        fileReader.readAsBinaryString(selectedFile);
        fileReader.onload = (event) => {
            let data = event.target.result;
            let workbook = XLSX.read(data, { type: "binary" });
            console.log(workbook, "workbook");
            workbook.SheetNames.forEach(sheet => {
                let rowObject = XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheet]);
                console.log(rowObject, "rowObject");
                let outputArr = [];
                let customers = [];
                let trackings = [];
                let obj = {}
                rowObject.forEach(d => {
                    let cust = rowObject.filter(l => l["Customer"] == d["Customer"])


                    cust.forEach(c => {
                        trackings.push(c["Tracking number"].trim());
                    })
                    console.log(trackings, "trackings")
                    if (!customers.includes(d["Customer"])) {
                        customers.push(d["Customer"]);
                        obj = {
                            "Customer": d["Customer"],
                            "Tracking number": trackings.join(",")
                        }
                        outputArr.push(obj)
                    }
                    trackings = []


                })
                exportWorksheet(rowObject, outputArr)
                document.getElementById("jsondata").innerHTML = JSON.stringify(rowObject, undefined, 4)
            });
        }
    }
});

window.onload = function () {
    document.querySelector("#exportWorksheet").click(function () {
        var josnData = document.querySelector("#josnData").value;
        var jsonDataObject = eval(josnData);
        exportWorksheet(jsonDataObject);
    });

    document.querySelector("#exportWorksheetPlus").click(function () {
        var josnData = document.querySelector("#josnData").value;
        var jsonDataObject = eval(josnData);
        exportWSPlus(jsonDataObject);
    });

};


function exportWorksheet(jsonObject, outputArr) {
    var myFile = "myFile.xlsx";
    var myWorkSheet = XLSX.utils.json_to_sheet(jsonObject);
    var output = XLSX.utils.json_to_sheet(outputArr);
    var myWorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(myWorkBook, myWorkSheet, "myWorkSheet");
    XLSX.utils.book_append_sheet(myWorkBook, output, "Output");
    XLSX.writeFile(myWorkBook, myFile);
}

function exportWSPlus(jsonObject) {
    var myFile = "myFilePlus.xlsx";
    var myWorkSheet = XLSX.utils.json_to_sheet(jsonObject);
    XLSX.utils.sheet_add_aoa(myWorkSheet, [["Your Mesage Goes Here"]], { origin: 0 });
    var merges = myWorkSheet['!merges'] = [{ s: 'A1', e: 'D1' }];
    var myWorkBook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(myWorkBook, myWorkSheet, "myWorkSheet");
    XLSX.writeFile(myWorkBook, myFile);
}

/*
[
    {
        "Customer": "Dingo Keina",
        "Tracking number": "1108361BAF2          "
    },
    {
        "Customer": "D Hourst",
        "Tracking number": "1108361BAE3          "
    },
    {
        "Customer": "Steve Newnham",
        "Tracking number": "1108361BAD4          "
    },
    {
        "Customer": "Tom Mortimore",
        "Tracking number": "1108361BAC8,1108361BABA,1108361BAAC,1108361BA9E"
    },
    {
        "Customer": "Mark Howarth",
        "Tracking number": "1108361B9F3          "
    },
    {
        "Customer": "Liam Watson",
        "Tracking number": "1108361B9E4          "
    },
    {
        "Customer": "B Langdon",
        "Tracking number": "1108361B9D5          "
    },
    {
        "Customer": "Tom Mortimore",
        "Tracking number": "1108361BAC9          "
    }
]
*/