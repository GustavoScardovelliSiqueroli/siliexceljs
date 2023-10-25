const express = require('express');
const exceljs = require('exceljs');
var fs = require('fs');

const app = express();

const PORT = 3000;

const path = require('path');
const { title } = require('process');
const router = express.Router();

router.get('/', (req, res) => {

    res.sendFile(path.join(__dirname + '/index.html'));
});

router.get('/export', async (req, res) => {
    try {
        let workbook = new exceljs.Workbook();

        const sheet = workbook.addWorksheet("casos")
        sheet.columns = [
            { header: "ID", key: "id", width: 25 },
            { header: "Endereço", key: "endereco", width: 25 },
            { header: "Status", key: "status", width: 25 },
        ];

        let object = JSON.parse(fs.readFileSync('cases.json', 'utf8'));

        await object.map((value, idx) => {
            let row = sheet.addRow({ id: value.id,
                     endereco: value.endereco,
                      status: value.status,
                     });
           
        });
        sheet.eachRow((row, rowNumber) => {
            row.eachCell((cell) => {
              cell.border = {
                top: { style: "thin" },
                left: { style: "thin" },
                bottom: { style: "thin" },
                right: { style: "thin" },
              };
            });
            row.commit();
            });

        for(i = 1; i <= 3; i++){
            sheet.getRow(1).getCell(i).fill = {
                type: "pattern",
                pattern: "solid",
                fgColor: { argb: "808080" },
              };
        }
        //Como faço para mudar somente a cor da letra e não a cor da célula inteira?
        for(i = 0; i <= sheet.rowCount; i++){
            if(sheet.getRow(i).getCell(3).value == "positivo"){
                sheet.getRow(i).getCell(3).font = {
                    type: 'pattern',
                    color: { argb: 'FF008000' } // ARGB value for green color
                }
            }
            if(sheet.getRow(i).getCell(3).value == "negativo"){
                sheet.getRow(i).getCell(3).font = {
                    type: 'pattern',
                    color: { argb: 'ff0000' } // ARGB value for green color
                }
            }
            if(sheet.getRow(i).getCell(3).value == "suspeito"){
                sheet.getRow(i).getCell(3).font = {
                    type: 'pattern',
                    color: { argb: 'FFA500' } // ARGB value for green color
                }
            }
            sheet.getRow(i).getCell(1).alignment = {
                vertical: 'middle',
                horizontal: 'center'
            }
        }
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.setHeader("Content-Disposition", "attachment; filename=" + "cases.xlsx");
        workbook.xlsx.write(res)

    } catch (error) {
        console.log(error);
    }
   

})

app.use('/', router);

app.listen(PORT, () => {
    console.log(`Server is runing on port: ${PORT}`);
});

