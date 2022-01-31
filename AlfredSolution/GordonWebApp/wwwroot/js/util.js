function alertAction(title, message, ok) {
    var modal = '<div class="modal fade" id="alertModalAction" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">';
    modal += '<div class="modal-dialog modal-dialog-centered" role="document">';
    modal += '<div class="modal-content">';
    modal += '<div class="modal-header">';
    modal += '<h5 class="modal-title" id="alertModalActionTitle">' + title + '</h5>';
    modal += '<button type="button" class="close" data-dismiss="modal" aria-label="Fechar">';
    modal += '<span aria-hidden="true">&times;</span>';
    modal += '</button>';
    modal += '</div>';
    modal += '<div id="alertModalActionBody" class="modal-body">';
    modal += message;
    modal += '</div>';
    modal += '<div class="modal-footer">';
    modal += '<button id="btnOkAlertModalAction" type="button" class="btn btn-primary">Ok</button>';
    modal += '<button type="button" class="btn btn-secondary" data-dismiss="modal">Cancelar</button>';
    modal += '</div>';
    modal += '</div>';
    modal += '</div>';
    modal += '</div>';

    $("body").append(modal);

    $('#alertModalActionBody').text(message);
    $('#alertModalActionTitle').text(title);
    $("#btnOkAlertModalAction").attr("onclick", ok);
    $('#alertModalAction').modal('show');
}

function alertModal(title, message, fechar) {
    var modal = '<div class="modal fade" id="alertModal" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">';
    modal += '<div class="modal-dialog modal-dialog-centered" role="document">';
    modal += '<div class="modal-content">';
    modal += '<div class="modal-header">';
    modal += '<h5 class="modal-title" id="alertModalTitle">' + title + '</h5>';
    modal += '<button type="button" class="close" data-dismiss="modal" aria-label="Fechar">';
    modal += '<span aria-hidden="true">&times;</span>';
    modal += '</button>';
    modal += '</div>';
    modal += '<div id="alertModalBody" class="modal-body">';
    modal += message;
    modal += '</div>';
    modal += '<div class="modal-footer">';
    modal += '<button type="button" id="btnFecharModal" class="btn btn-secondary" data-dismiss="modal">Fechar</button>';
    modal += '</div>';
    modal += '</div>';
    modal += '</div>';
    modal += '</div>';

    $("body").append(modal);

    $('#alertModalBody').text(message);
    $('#alertModalTitle').text(title);
    $("#btnFecharModal").attr("onclick", fechar);
    $('#alertModal').modal('show');
}

function atualizarStatus(id, ativar) {

    var tdBtn = $("#btn" + id);
    var tdAtivo = $("#at" + id);

    var urlAlteracao = "/Usuario/Ativar"
    if (!ativar) {
        urlAlteracao = "/Usuario/Desativar"
    }

    var conteudoAt = "Não";
    if (ativar)
        conteudoAt = "Sim";

    var conteudoBtn = '<button id="editar" type="button" class="btn btn-w-m btn-default" onclick="alterarUsuario(' + id + ')">Editar</button> ';

    if (!ativar)
        conteudoBtn += '<button id="' + id + '" type="button" class="btn btn-w-m btn-primary" onclick="ativarUsuario(' + id + ')"><span id="spn' + id +'" class="spinner-border spinner-border-sm d-none" role="status" ></span>Ativar</button> ';
    else
        conteudoBtn += '<button id="' + id + '" type="button" class="btn btn-w-m btn-default" onclick="desativarUsuario(' + id + ')"><span id="spn' + id +'" class="spinner-border spinner-border-sm d-none" role="status" ></span>Desativar</button>';


    $.ajax({
        type: 'GET',
        url: urlAlteracao,
        dataType: "JSON",
        cache: false,
        async: true,
        data: { "id": id },
        success: function (data) {
            if (data.sucess) {
                tdBtn.html(conteudoBtn);
                tdAtivo.text(conteudoAt)
                alertModal("Alerta", data.message);
            }
        }
    });
}

function spinnerButton(idButton, textButton, ativar) {
    if (ativar) {
        $("#" + idButton).prop('disabled', true);
        $("#" + idButton).empty().html('<span class="spinner-border spinner-border-sm" role="status" ></span> ' + textButton);
    } else {
        $("#" + idButton).prop('disabled', false);
        $("#" + idButton).empty().text(ativar);
    }
}