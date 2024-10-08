function getFileSize(size) {
    if (size > 1073741824) { return `${(size / 1073741824).toFixed(2)}GB`; }
    if (size > 1048576) { return `${(size / 1048576).toFixed(2)}MB`; }
    if (size > 1024) { return `${(size / 1024).toFixed(2)}KB`; }
    return `${size}B`;
}
function drapDropFiles(ext = "") {
    var uploadArea = $('#uploadfile'); if (uploadArea.length == 0) { return; } /** Nếu không tồn tại thì không áp dụng */
    var uploadButton = $('#uploadButton');
    ext = ext.trim();
    var exts = ext.replace(" ", "").toLowerCase().split(',').filter(extension => extension !== ''); /* .xlsx,.xls,.db,.zip, ... */
    var fileList = [];
    /* Ngăn chặn hành vi mặc định của trình duyệt */
    uploadArea.on('dragenter dragover', function (e) { e.stopPropagation(); e.preventDefault(); uploadArea.addClass('hover'); });
    uploadArea.on('dragleave', function (e) { e.stopPropagation(); e.preventDefault(); uploadArea.removeClass('hover'); });
    uploadArea.on('drop', function (e) {
        e.stopPropagation();
        e.preventDefault();
        uploadArea.removeClass('hover');
        var files = e.originalEvent.dataTransfer.files;
        handleFiles(files);
    });
    /* Nhấp để chọn file */
    uploadArea.on('click', function () { $('<input type="file" multiple>').on('change', function (e) { var files = e.target.files; handleFiles(files); }).click(); });
    /* Hàm xử lý file và hiển thị thông tin */
    function handleFiles(files) {
        $('#fileList').empty();
        fileList = [];
        for (var i = 0; i < files.length; i++) {
            var file = files[i];
            if (ext != "") {
                if (!exts.some(extension => file.name.toLowerCase().endsWith(extension))) { continue; }
            }            
            var fileSize = getFileSize(file.size);
            var listItem = $('<li class="list-group-item d-flex justify-content-between align-items-center"></li>');
            var fileInfo = $('<span></span>').text(`${file.name} (${fileSize})`);
            /* Kiểm tra nếu là hình ảnh */
            if (file.type.startsWith('image/')) {
                var reader = new FileReader();
                reader.onload = (function (fileInfo) {
                    return function (e) { var img = $('<img>').attr('src', e.target.result); listItem.prepend(img); };
                })(fileInfo);
                reader.readAsDataURL(file);
            }
            listItem.append(fileInfo);
            $('#fileList').append(listItem);
            fileList.push(file); /* Thêm file vào danh sách */
        }
        /* Kích hoạt nút upload nếu có file */
        if (fileList.length > 0) { uploadButton.prop('disabled', false); }
        else { uploadButton.prop('disabled', true); }
    }

    /* Xử lý sự kiện khi nhấn nút upload */
    uploadButton.on('click', function () {
        if (fileList.length === 0) { messageBox("JS Thông báo", "Vui lòng chọn ít nhất một file để tải lên."); return; }
        var formData = new FormData();
        for (var i = 0; i < fileList.length; i++) { formData.append('files[]', fileList[i]); }
        var url = uploadArea.attr("data-urlpost");
        var target = uploadArea.attr("data-target");
        postform(formData, url, target);
    });
}