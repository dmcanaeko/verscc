var formParams = {
    fileBlob: '',
    inputFileID: '$("#custom_fileupload_holder")',
    randomNumber: '',
    allowedFileType: '',
    maxFileSize: '4000000',
    maxFileSizeDisplay: '4000000',
    imgClickSelector: '',
    deleteFileSelector: '',
    kdfSaveFlag: false,
    full_classification: '',
    fileUploadUrl: 'https://graph.microsoft.com/v1.0/sites/sccextranet.sharepoint.com/drive/items/',
    fieldNames: []
}

function setFiledNames(fieldnames) {
    formParams.fieldNames = fieldnames;
}

function setFileUploadUrl(url) {
    formParams.fileUploadUrl = url;
}

function do_KDF_Ready_Sharepoint(event, kdf) {
    var template_name = KDF.getVal('txt_FT_template');
    if (KDF.getVal('txt_FT_template') == '' || $('#dform_widget_txt_FT_template').length < 1) {
        template_name = 'FT_template1';
    }
    KDF.customdata('sharepoint_config', '', true, true, {
        ft_operation: 'file_list',
        txt_FT_template: template_name
    })

    if ((KDF.kdf().form.readonly && KDF.kdf().access == 'citizen') || (KDF.kdf().viewmode == 'R')) {
        KDF.makeReadonly();
        KDF.hideWidget('ahtm_custom_fileupload');
        KDF.showSection('area_file_view_mode');
        KDF.customdata('sharepoint_token', 'imitateKdfReady readonly', true, true, {});
    } else if (KDF.kdf().viewmode == 'U') {
        KDF.showWidget('ahtm_custom_fileupload');
        KDF.hideSection('area_file_view_mode');
        KDF.customdata('sharepoint_token', 'imitateKdfReady readonly', true, true, {});
    }

    var CustomFileUploadWidget = $('#custom_fileupload_holder');
    $(document).on('drop dragover', function(e) {
        e.preventDefault();
    });
    $(document).on('change', '#custom_fileupload', function() {
        var fileName = $("#custom_fileupload")[0].files[0].name;
        var fileNameClean = fileName.split('.').pop();
        var template_name = KDF.getVal('txt_FT_template');

        if (KDF.getVal('txt_FT_template') == '' || $('#dform_widget_txt_FT_template').length < 1) {
            template_name = 'FT_template1';
        }
        KDF.customdata('sharepoint_config', '', true, true, {
            txt_FT_template: template_name,
            txt_file_format: fileNameClean
        })

    });

    $('body').on('click', 'img', function() {
        if ($(this).data('fieldname')) {
            if (KDF.kdf().form.readonly) {
                formParams.imgClickSelector = $(this).data('fieldname');
                KDF.customdata('sharepoint_token', 'imgClickEvent', true, true, {});
            }
        }
    })

    $('body').on('click', '.fileicon', function() {
        if ($(this).data('fieldname')) {
            if (KDF.kdf().form.readonly) {
                formParams.imgClickSelector = $(this).data('fieldname');
                KDF.customdata('sharepoint_token', 'imgClickEvent', true, true, {});
            }
        }
    })

    $('body').on('click', '.delete_file', function() {
        formParams.deleteFileSelector = $(this).closest('span').data('fieldname');
        KDF.customdata('sharepoint_token', 'imgClickEvent', true, true, {});
    })
}

function setFileBlobData(fileBlob) {
    formParams.fileBlob = fileBlob;
}

function processFile() {
    var fileError = false;
    var fileName = $("#custom_fileupload")[0].files[0].name;
    var fileNameClean = fileName.split('.').pop();

    if ($("#custom_fileupload")[0].files[0].size <= formParams.maxFileSize) {
        fileError = false;
    } else {
        fileError = true;
        KDF.showError('File size is too large');
    }

    if (!fileError) {
        fileNames = [];
        if (formParams.fieldNames.every(function(fieldName) { return KDF.getVal('txt_filename_' + fieldName) !== '' })) {
            fileError = true;
            KDF.showError('Maximum number of file uploads has been reached');
        }
    }

    if (!fileError) {
        for (var i = 0; i < formParams.fieldNames.length; i++) {
            if (KDF.getVal('txt_filename_' + formParams.fieldNames[i]) == fileName) {
                fileError = true;
                KDF.showError('A file with this name already exists');
                break;
            }
        }
    }

    if (!fileError) {
        KDF.hideMessages();
        $(".dform_fileupload_progressbar").html("<div style='width: 0%;'>");
        var selector = formParams.inputFileID;

        $(".dform_fileupload_progressbar").html("<div style='width: 10%;'>");

        $("#custom_fileupload").prop('disabled', true);

        var reader = new FileReader();
        reader.readAsArrayBuffer($("#custom_fileupload")[0].files[0]);

        reader.onloadend = function() {
            setFileBlobData(reader.result);
            $(".dform_fileupload_progressbar").html("<div style='width: 30%;'>");
            if (!formParams.kdfSaveFlag) {
                KDF.save();
                document.getElementById("custom_fileupload_holder").focus();
            } else {
                KDF.customdata('sharepoint_token', 'imitateKdfReady', true, true, {});
            }
        };
    }

}

function setFileThumbnails(access_token) {
    formParams.fieldNames.forEach(function(name) {
        if (KDF.getVal('txt_filename_' + name) !== '') {
            sharepointFileThumbnail(KDF.getVal('txt_sharepointID_' + name), access_token, 'txt_filename_' + name, name);
        }
    });
}

function do_KDF_Custom_Sharepoint(response, action) {
    if (action === 'sharepoint_token') {
        var access_token = response.data['access_token'];
        if (!KDF.kdf().form.readonly && formParams.deleteFileSelector == '') {

            if (KDF.kdf().viewmode == 'U' && formParams.fileBlob == '') {
                setFileThumbnails(access_token);
            } else if (formParams.fileBlob !== '') {

                if (!formParams.kdfSaveFlag) {
                    formParams.kdfSaveFlag = true;
                    formParams.full_classification = response.data['full_classification'];
                }

                sharepointFileUploader(access_token);
            }


        } else if (!KDF.kdf().form.readonly && formParams.deleteFileSelector !== '') {
            deleteFile(access_token);
        }

        if (KDF.kdf().form.readonly && formParams.imgClickSelector == '') {
            //sharepointFileThumbnail (itemID, access_token)
            setFileThumbnails(access_token);
        } else if (KDF.kdf().form.readonly && formParams.imgClickSelector !== '') {
            sharepointDownloadFile(access_token)
        }
    } else if (action == 'sharepoint_config') {
        if (response.data['pass_status']) {
            // processFile();

            if (response.data['pass_status'] == 'good') {
                processFile();
            } else {
                KDF.showError('Incorrect file type selected.')
            }
        } else {
            var sharepoint_title = '';
            if ($('#dform_widget_txt_sharepoint_title').length > 0) {
                sharepoint_title = KDF.getVal('txt_sharepoint_title');

            } else {
                sharepoint_title = 'Please upload files';
            }
            var txt_file_types = response.data['txt_file_types'];
            formParams.allowedFileType = txt_file_types.replace(/'/g, '').replace('(', '').replace(')', '').replace(/,/g, ', ');
            formParams.maxFileSizeDisplay = response.data['txt_max_filesize'];

            if ($('#custom_fileupload_holder').length > 0) {
                var widget = '<div data-type="file" data-name="file_ootb" data-active="true" data-agentonly="false" class="file-progress">' +
                    '<div><label>' + sharepoint_title + '</div></label>' +
                    '<div style="position: relative;"><input id="custom_fileupload" type="file" name="uploadedFile" aria-label="Upload file">' +
                    '<span class="file-gov-icon"><span class="file-gov-icon-a"></span><span class="file-gov-icon-b"></span><label class="file-gov-text">Upload file</label></span>' +
                    '<div class="helptext">Accepted file types are ' + formParams.allowedFileType + ' up to ' + formParams.maxFileSizeDisplay + ' MB in size</div>' +
                    '<div class="dform_fileupload_progressbar" id="custom_fileupload_progressbar"></div>' +
                    '<div class="filenames" id="custom_fileupload_files"></div><br><br></div>' +
                    ' </div>';

                $('#custom_fileupload_holder').html(widget);
            }
        }
    }
}

function do_KDF_Save_Sharepoint() {

    if (formParams.fileBlob !== '') {
        $('#custom_fileupload').focus();
    }

    if (!formParams.kdfSaveFlag) {
        if (formParams.fileBlob !== '') {
            $('#custom_fileupload').focus();
            $('#dform_successMessage').remove();
            //formParams.kdfSaveFlag = true;
            KDF.customdata('sharepoint_token', 'imitateKdfReady', true, true, { 'SaveForm': 'true', 'caseid': KDF.kdf().form.caseid });
        }
    }
}

function sharepointFileUploader(access_token) {
    KDF.lock();
    var fileName = $("#custom_fileupload")[0].files[0].name;
    var fileSize = $("#custom_fileupload")[0].files[0].size;
    var uploadURL = formParams.fileUploadUrl + 'root:/Verint/' + formParams.full_classification + '/' + KDF.kdf().form.caseid + '/' + fileName + ':/content';

    $(".dform_fileupload_progressbar").html("<div style='width: 50%;'>");
    $.ajax({
        url: uploadURL,
        dataType: 'json',
        processData: false,
        headers: { 'Authorization': access_token },
        data: formParams.fileBlob,
        method: 'PUT',

    }).done(function(response) {
        sharepointFileThumbnail(response.id, access_token)
        $(".dform_fileupload_progressbar").html("<div style='width: 60%;'>");

        for (var i = 0; i < formParams.fieldNames.length; i++) {
            var name = formParams.fieldNames[i];
            if (KDF.getVal('txt_sharepointID_' + name) == '') {
                KDF.setVal('txt_sharepointID_' + name, response.id);
                KDF.setVal('txt_filename_' + name, fileName);
                KDF.setVal('txt_sharepoint_link_' + name, response['webUrl']);
                break;
            }
        }

        KDF.save();
    });


}

function sharepointFileThumbnail(itemID, access_token, widgetName, fieldName) {
    var getThumbnailURL = formParams.fileUploadUrl + itemID + '/thumbnails';

    $.ajax({
        url: getThumbnailURL,
        dataType: 'json',
        headers: { Authorization: access_token },
        method: 'GET',

    }).done(function(response) {
        var thumbnailURL = (response.value[0]) ? response.value[0].medium['url'] : undefined;
        if (!KDF.kdf().form.readonly) {
            if (KDF.kdf().viewmode === 'U' && formParams.fileBlob == '') {
                if (fieldName) {
                    KDF.setVal('txt_filename_' + fieldName + '_thumb', thumbnailURL);
                }
                addFileContainer(fieldName);
            } else if (formParams.fileBlob !== '') {

                $(".dform_fileupload_progressbar").html("<div style='width: 60%;'>");

                for (var i = 0; i < formParams.fieldNames.length; i++) {
                    var name = formParams.fieldNames[i];
                    if (KDF.getVal('txt_filename_' + name + '_thumb') == '') {
                        KDF.setVal('txt_filename_' + name + '_thumb', thumbnailURL);
                        break;
                    }
                }
                $(".dform_fileupload_progressbar").html("<div style='width: 100%;'>");
                setTimeout(function() {
                    addFileContainer();
                    $(".dform_fileupload_progressbar").html("<div style='width: 0%;'>");
                }, 1000);
            }

        } else if (KDF.kdf().form.readonly || KDF.kdf().viewmode == 'R') {
            var fileName = KDF.getVal(widgetName);
            var html;

            html = '<div id="' + widgetName + '"style="float: left;">' +
                '<div style="margin-right: 10px">' + getImage(thumbnailURL, widgetName, fileName, fieldName) +
                '</div><div>' + fileName + '</div></div>';

            setTimeout(function() { $('#custom_fileupload_view').append(html) }, 1000);
        }
    });

    $("#custom_fileupload").prop('disabled', false);


}

function addFileContainer(fieldName) {
    $('input#custom_fileupload').val('');
    var fileName;
    var fileThumbnail;
    var widgetName = 'txt_filename_' + fieldName;

    if (KDF.kdf().viewmode == 'U' && formParams.fileBlob == '') {
        fileName = KDF.getVal(widgetName);
        fileThumbnail = KDF.getVal(widgetName + '_thumb');
    } else if (formParams.fileBlob !== '') {
        for (var i = 0; i < formParams.fieldNames.length; i++) {
            fieldName = formParams.fieldNames[i];
            if ($('.filenames .txt_filename_' + fieldName).length < 1) {
                fileName = KDF.getVal('txt_filename_' + fieldName);
                fileThumbnail = KDF.getVal('txt_filename_' + fieldName + '_thumb');
                widgetName = 'txt_filename_' + fieldName;
                break;
            }
        }
    }

    $(".filenames").append('<span class="' + widgetName + '"> <span class="img_container"> ' + getImage(fileThumbnail, widgetName, fileName, fieldName) +
        '<div>' + fileName + '<span id="delete_' + widgetName + '" data-fieldname="' + fieldName + '" style="font-weight:bold;" class="delete_file">4</span></div></span></span>');

    KDF.unlock();
}

function getImage(fileThumbnail, widgetName, fileName, fieldName) {
    if (fileThumbnail) {
        return ' <img id="img_' + widgetName + '" data-filename="' + fileName + '"data-fieldname="' + fieldName + '" src=' + fileThumbnail + '" > ';
    } else {
        return ' <span class="fileicon" data-filename="' + fileName + '" data-fieldname="' + fieldName + '">e</span> ';
    }
}


function sharepointDownloadFile(access_token) {
    // var selector = formParams.imgClickSelector;
    // var fieldName = $('.' + selector).data('fieldname');
    var sharepointID = KDF.getVal('txt_sharepointID_' + formParams.imgClickSelector);
    var getFileURL = formParams.fileUploadUrl + sharepointID + '/preview';

    $.ajax({
        url: getFileURL,
        headers: { Authorization: access_token },
        type: 'POST'
    }).done(function(response) {
        window.open(response.getUrl);
    }).fail(function() {});
    formParams.imgClickSelector = '';
}

function deleteFile(access_token) {
    $(".dform_fileupload_progressbar").html("<div style='width: 0%;'>");
    var selector = formParams.deleteFileSelector;
    var fileID = KDF.getVal('txt_sharepointID_' + selector);
    var deleteURL = formParams.fileUploadUrl + fileID;
    $.ajax({
        url: deleteURL,
        processData: false,
        headers: { 'Authorization': access_token },
        method: 'DELETE'

    }).done(function(response) {
        $('span.txt_filename_' + selector).remove();
        KDF.setVal('txt_sharepointID_' + selector, '')
        KDF.setVal('txt_filename_' + selector, '')
        KDF.setVal('txt_filename_' + selector + '_thumb', '')
        KDF.save();
    }).fail(function() {
        KDF.showError('Delete file has failed, please try again');
    });

    formParams.deleteFileSelector = '';
}