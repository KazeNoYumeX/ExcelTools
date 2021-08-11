// Get a reference to the file input element
const inputElement = document.querySelector('input[type="file"]');

// Create the FilePond instance
const pond = FilePond.create(inputElement, {
    allowMultiple: true,
    allowReorder: true,
});

// Easy console access for testing purposes
window.pond = pond;

const extensionList = ['xlsx', 'xls', 'xml']
let fileName = ''

$('#excel-file').change((e) => {
    const files = e.target.files;
    const extension = files[0].name.split('.').pop()
    let excelData

    if (extensionList.indexOf(extension) !== -1) {
        let fileReader = new FileReader();
        fileReader.onload = (ev)=> {
            try {
                const data = ev.target.result
                // 以二進位制流方式讀取得到整份excel表格物件
                const workbook = XLSX.read(data, {
                    type: 'binary'
                })
                // 儲存獲取到的資料
                let excelData = [];
            } catch (e) {
                console.log('檔案型別不正確');
                return;
            }
            // 表格的表格範圍，可用於判斷表頭是否數量是否正確
            let fromTo = '';
            // 遍歷每張表讀取
            for (let sheet in workbook.Sheets) {
                if (workbook.Sheets.hasOwnProperty(sheet)) {
                    fromTo = workbook.Sheets[sheet]['!ref'];
                    excelData = excelData.concat(XLSX.utils.sheet_to_json(workbook.Sheets[sheet]));
                    //  break; // 如果只取第一張表，就取消註釋這行
                }
            }
            fileName = excelData
        };
        // 以二進位制方式開啟檔案
        fileReader.readAsBinaryString(files[0]);
    }
});

const formatConversion = (name, serial) => {
    const utc_days = Math.floor(serial - 25569);
    const utc_value = utc_days * 86400;
    const date_info = new Date(utc_value * 1000);

    name = name.replace(/\s*/g, "").toLowerCase()
    const nameYear = date_info.getFullYear()
    const password = name.substr(0, 1) + date_info.getMonth() + date_info.getDate() + date_info.getMonth() + date_info.getDate()
    name = name + nameYear

    console.log(name, password)

    return {
        name,
        password
    };
}

const onExport = () => {
    for (let i = 0; i < fileName.length; i++) {
        fileName[i] = formatConversion(fileName[i].name, fileName[i].date)
    }

    const workbook = XLSX.utils.book_new();
    const ws1 = XLSX.utils.json_to_sheet(fileName);

    XLSX.utils.book_append_sheet(workbook, ws1, "Sheet1");

    const wbout = XLSX.write(workbook, {
        bookType: 'xlsx',
        bookSST: true,
        type: 'binary'
    });
    const wboutBin64 = btoa(wbout);
    const a = document.createElement('a');
    a.href = 'data:;base64,' + wboutBin64;
    a.download = '基本資料表.xlsx';
    a.click();
}