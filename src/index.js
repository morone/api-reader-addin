/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as OfficeHelpers from '@microsoft/office-js-helpers';

$(document).ready(() => {
    $('#run').click(run);
});
  
// The initialize function must be run each time a new page is loaded
Office.initialize = (reason) => {
    $('#sideload-msg').hide();
    $('#app-body').show();
};

async function run(){
    try{

        await Excel.run(async context => {

            var columns = ("ABCDEFGHIJKLMNOPQRSTUVWXYZ").split("");
            
            var result;
            var keys;
            var colIndex = 0;
            var rowIndex = 2;
            var colNames = [];
            var self = this;
            var sheet = context.workbook.worksheets.getActiveWorksheet();

            sheet.getRange().clear();
            
            await context.sync();

            jQuery.ajax({
                context: this,
                url: $('#api-url').val(),
                type: "GET",
                contentType: 'application/json',
                timeout: 120000,
                async: false,
            }).done(function(r){
                this.result = r;
                this.keys = Object.keys(r[0]);
            });

            

            this.keys.forEach(function(header){
                sheet.getRange(columns[colIndex] + "1").values = [[ header ]];
                colNames.push(header);
                colIndex++;
            });

            colIndex = 0;
            
            console.log(colNames);

            this.result.forEach(function(data){
                self.keys.forEach(function(header){
                    sheet.getRange(columns[colIndex] + rowIndex).values = [[ JSON.stringify(data[colNames[colIndex]]) ]];
                    colIndex++;                    
                }, this);
                colIndex = 0;
                rowIndex++;
            });

        });

    }catch(error) {
        OfficeHelpers.UI.notify(error);
        OfficeHelpers.Utilities.log(error);
    };
}
