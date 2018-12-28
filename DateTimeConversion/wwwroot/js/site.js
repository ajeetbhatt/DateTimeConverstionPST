var fileUpload = {
    excelImport: {
        _validFilesTypes: ["xls", "xlsx"],
    
        processImport: function () {
            var input = document.getElementById('excelFileUpload');
            var files = input.files;
            var formData = new FormData();
            for (var i = 0; i < files.length; i++) {
                formData.append("files", files[i]);
            }
            waitingDialog.show("Please wait while converting date and time");
            $.ajax({
                url: "/Home/UploadFile",
                data: formData,
                processData: false,
                contentType: false,
                type: "POST",
                success: function (response) {
                    waitingDialog.hide();
                 $("#downloadBtn")[0].click();
             
                },
                error: function () {
                    waitingDialog.hide();
                    $.notify("There are some server side error while processing request", { position: "top right", className: "error" });
                }
            });
        },
        validateFileUpload: function (file) {

            var isValidFile = false;
            var filePath = file.value;
            if (typeof (filePath) !== 'undefined') {
                var ext = filePath.substring(filePath.lastIndexOf('.') + 1).toLowerCase();

                for (var i = 0; i < fileUpload.excelImport._validFilesTypes.length; i++) {
                    if (ext === fileUpload.excelImport._validFilesTypes[i]) {
                        isValidFile = true;
                        $('#BtnUploadExcelFile').removeAttr('disabled');
                        break;
                    }
                }

                if (!isValidFile) {
                    file.value = null;
                    $('#BtnUploadExcelFile').attr('disabled', 'disabled');
                    $("#importModal").modal('show');
                }
            }

            return isValidFile;
        }

    }
}