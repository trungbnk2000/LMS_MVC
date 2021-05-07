CREATE
$(function () {
    var PlaceHolderElement = $('#PlaceHolderHere');
    $('button[data-toggle="ajax-modal"]').click(function (event) {
        var url = $(this).data('url'); //url = createmodal

        $.get(url).done(function (data) {
            PlaceHolderElement.html(data)
            PlaceHolderElement.find('.modal').modal('show');
        })
    })

    $('#btnSubmit').click(function () {
        var myformdata = $('#myForm').serialize();
    
        $.ajax({
            type: "POST",
            url: "/Books/CreateModal",
            data: myformdata,
            success: function () {
                $('#addBook').modal("hide");
            }
        })
    })
})

