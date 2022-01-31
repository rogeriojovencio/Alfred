
$(document).on("click", "#btnUpload", function (e) {

    var botao = $(this);
    $(botao).ladda();
    $(botao).ladda('start');
    var formData = new FormData(document.getElementById("FormImportacao"));
    $.ajax({
        type: "POST",
        url: "/Importacao/ImportacaoUploadFile/",
        data: formData,
        processData: false,
        contentType: false,
        dataType: "json",
        success: function (data) {
            $(botao).ladda('stop');
            if (data.isValid) {
                swal(data.title, data.message, data.type).then(function () { window.location.reload(); });

            } else {
                swal(data.title, data.message, data.type);
            }
        },
        error: function (data) {
            $(botao).ladda('stop');
        }
    });
});



$(document).on("click", ".btn-excluir", function (e) {
    var botao = $(this);
    $(botao).ladda();
    $(botao).ladda('start');


    var model = {
        id: botao.attr("data-identity"),
        __RequestVerificationToken: $('input[name="__RequestVerificationToken"]').val(),
    }

    $.ajax({
        type: "POST",
        url: "/Importacao/ExcluirImportacao/",
        data: JSON.stringify( model),
        async: true,
        data: model,
        datatype: "html",
        success: function (data) {
            $(botao).ladda('stop');
            if (data.IsValid) {
                swal(data.title, data.message, data.type);
                Status();
                $('#fileUpload').val('');

            } else {
                swal(data.title, data.message, data.type);
            }
        },
        error: function (data) {
            $(botao).ladda('stop');
        }
    });
});


$(document).ready(function () {
    Status();
    var myInterval = setInterval(function () { Status() }, 5000);
});

function Status() {

    $.ajax({
        type: "GET",
        url: "/Importacao/StatusImportacao/",
        async: true,
        datatype: "html",
        success: function (Status) {
            $('#contentStatusImportacao').html(Status);
        }
    });
}
