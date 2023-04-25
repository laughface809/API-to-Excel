import Excel from 'exceljs';
import path from 'path';
import axios from 'axios';
import { ids } from './id';

const exportCountriesFile = async () => {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet('Countries List');

  worksheet.columns = [
    { key: 'quantity', header: 'Stock' },
    { key: 'description', header: 'Description' },
    { key: 'productCode', header: 'SKU' },
    { key: 'category', header: 'Category' },
  ];

  worksheet.columns.forEach((sheetColumn) => {
    sheetColumn.font = {
      size: 12,
    };
    sheetColumn.width = 30;
  });

  worksheet.getRow(1).font = {
    bold: true,
    size: 13,
  };

  ids.forEach((item, index) => {
    /**
     * @param {any} this looping is only for get how much id we have
     * item is to get data ID from object and index is using for throtle to make timeout request
     * so the request is not refused by the server, usually because of firewall or load balancing
     * 
     * https://stackoverflow.com/questions/36706861/delay-between-each-iteration-of-foreach-loop
     */
    setTimeout(() => {
      axios
        .get(`{Secret URL}/${item.ID}`)
        .then(function (response) {
          const resData: any = response.data.data;

          /**
           * in here we will have much validation to check if the data available or not
           * at the end we change it to string
           * 
           * https://stackoverflow.com/questions/26732123/turn-properties-of-object-into-a-comma-separated-list
           */
          worksheet.addRow({
            quantity: (resData.productVariations
              ? resData.productVariations.length > 0 &&
                resData.productVariations.map(() => {
                  try {
                    return resData.productVariations[0].inventories[0].quantity;
                  } catch (e) {
                    return 0;
                  }
                })
              : 0
            ).toString(),
            description: resData.description,
            productCode: resData.productCode,
            category: (resData.vendorCategories
              ? resData.vendorCategories.length > 0 &&
                resData.vendorCategories
                  .map((item: { name: any }) => {
                    try {
                      return item.name;
                    } catch (e) {
                      return '';
                    }
                  })
                  .join(',')
              : ''
            ).toString(),
          });

          /**
           * we will call exceljs function after we get the response and write the file to xlsx format
           * if no  file, the file will generate automatically but if there, the file will overwrite
           */

          const exportPath = path.resolve(__dirname, 'getDetail.xlsx');

          workbook.xlsx.writeFile(exportPath);
        })
        .catch(function (error) {
          console.log(error);
        })
        .finally(function () {
          // always executed
        });
    }, 300 * index);
  });
};

exportCountriesFile();
