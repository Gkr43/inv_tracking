sap.ui.define([
    "sap/m/MessageToast",
    "sap/ui/core/Fragment",
    "invtracking/lib/xlsx"
], function (MessageToast, Fragment, xlsx) {
    'use strict';

    return {
        excelSheetsData: [],
        pDialog: null,
        UploadAcknowledgement: function (oEvent) {
            var oView = this.getView();

            if (!this.pDialog) {
                Fragment.load({
                    id: "excel_upload",
                    name: "invtracking.ext.fragment.ExcelUpload",
                    type: "XML",
                    controller: this
                }).then((oDialog) => {
                    var oFileUploader = Fragment.byId("excel_upload", "uploadSet");
                    oFileUploader.removeAllItems();
                    this.pDialog = oDialog;
                    this.pDialog.open();
                })
                    .catch(error => alert(error.message));
            } else {
                var oFileUploader = Fragment.byId("excel_upload", "uploadSet");
                oFileUploader.removeAllItems();
                this.pDialog.open();
            }
        },
        onUploadSet: function (oEvent) {
            console.log("Upload Button Clicked!!!")
            // checking if excel file contains data or not
            if (!this.excelSheetsData.length) {
                MessageToast.show("Select file to Upload");
                return;
            }

            var that = this;
            var oSource = oEvent.getSource();

            // creating a promise as the extension api accepts odata call in form of promise only
            var fnAddMessage = function () {
                return new Promise((fnResolve, fnReject) => {
                    that.callOdata(fnResolve, fnReject);
                });
            };

            var mParameters = {
                sActionLabel: oSource.getText() // or "Your custom text" 
            };
            // calling the oData service using extension api
            this.extensionAPI.securedExecution(fnAddMessage, mParameters);

            this.pDialog.close();
            /* TODO:Call to OData */
        },
        onTempDownload: function (oEvent) {
            console.log("Template Download Button Clicked!!!")
            var oModel = this.getView().getModel();
            // get the property list of the entity for which we need to download the template
            var oBuilding = oModel.getServiceMetadata().dataServices.schema[0].entityType.find(x => x.name === 'ZINVOICE_TRACKING_PType');
            // set the list of entity property, that has to be present in excel file template
            var propertyList = ['Vbeln', 'Acknumber','Vehno','Dname','Kunnr','Cname','Unlod','Lrnum'];

            var excelColumnList = [];
            var colList = {};

            // finding the property description corresponding to the property id
            propertyList.forEach((value, index) => {
                let property = oBuilding.property.find(x => x.name === value);
                colList[property.name] = '';
            });
            excelColumnList.push(colList);

            // initialising the excel work sheet
            const ws = XLSX.utils.json_to_sheet(excelColumnList);
            // creating the new excel work book
            const wb = XLSX.utils.book_new();
            // set the file value
            XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');
            // download the created excel file
            XLSX.writeFile(wb, 'Acknowledgement Template.xlsx');

            MessageToast.show("Template File Downloading...");
            /* TODO: Excel file template download */
        },
        onCloseDialog: function (oEvent) {
            this.pDialog.close();
        },
        onBeforeUploadStart: function (oEvent) {
            console.log("File Before Upload Event Fired!!!")
            /* TODO: check for file upload count */
        },
        onUploadSetComplete: function (oEvent) {
            console.log("File Uploaded!!!")
            // getting the UploadSet Control reference
            var oFileUploader = Fragment.byId("excel_upload", "uploadSet");
            // since we will be uploading only 1 file so reading the first file object
            var oFile = oFileUploader.getItems()[0].getFileObject();

            var reader = new FileReader();
            var that = this;

            reader.onload = (e) => {
                // getting the binary excel file content
                let xlsx_content = e.currentTarget.result;

                let workbook = XLSX.read(xlsx_content, { type: 'binary' });
                // here reading only the excel file sheet- Sheet1
                var excelData = XLSX.utils.sheet_to_row_object_array(workbook.Sheets["Sheet1"]);
                that.excelSheetsData= [];
                workbook.SheetNames.forEach(function (sheetName) {
                    // appending the excel file data to the global variable
                    that.excelSheetsData.push(XLSX.utils.sheet_to_row_object_array(workbook.Sheets[sheetName]));
                });
                console.log("Excel Data", excelData);
                console.log("Excel Sheets Data", this.excelSheetsData);
            };
            reader.readAsBinaryString(oFile);

            MessageToast.show("Upload Successful");
            /* TODO: Read excel file data*/
        },
        onItemRemoved: function (oEvent) {
            console.log("File Remove/delete Event Fired!!!");
            this.excelSheetsData = [];
            /* TODO: Clear the already read excel file data */
        },
        callOdata: function (fnResolve, fnReject) {
            //  intializing the message manager for displaying the odata response messages
            var oModel = this.getView().getModel();

            // creating odata payload object for Building entity
            var payload = {};

            this.excelSheetsData[0].forEach((value, index) => {
                // setting the payload data
                payload = {
                    "Vbeln": value["Vbeln"].toString(),
                    "Acknumber": value["Acknumber"].toString(),
                   "Vehno": value["Vehno"].toString(),
                    "Dname": value["Dname"].toString(),
                   "Kunnr": value["Kunnr"].toString(),
                    "Cname": value["Cname"].toString(),
                    "Unlod": value["Unlod"].toString(),
                    "Lrnum": value["Lrnum"].toString()  
                };
                // setting excel file row number for identifying the exact row in case of error or success
                //payload.ExcelRowNumber = (index + 1);
                // calling the odata service
                oModel.create("/ZINVOICE_TRACKING_P", payload, {
                    success: (result) => {
                        console.log(result);
                        var oMessageManager = sap.ui.getCore().getMessageManager();
                        var oMessage = new sap.ui.core.message.Message({
                            message: "Acknowledgement with ID: " + result.Acknumber + " updated Successfully",
                            persistent: true, // create message as transition message
                            type: sap.ui.core.MessageType.Success
                        });
                        oMessageManager.addMessages(oMessage);
                        fnResolve();
                    },
                    error: fnReject
                });
            });
        }
    };
});