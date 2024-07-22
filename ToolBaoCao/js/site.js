var alertkey = { danger: 'danger', warning: 'warning', info: 'info', success: 'success' }
var idmsg = '#modal-message';
function LTrim(value) { return value.replace(/\s*((\S+\s*)*)/, "$1"); }
function RTrim(value) { return value.replace(/((\s*\S+)*)\s*/, "$1"); }
function trim(value) { return LTrim(RTrim(value)); }
function check_email(email_id) {
    emailRegExp = /^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.([a-z]){2,4})$/;
    if (emailRegExp.test(document.getElementById(email_id).value)) { return true; }
    return false;
}
function textCounter(inputcheck, inputcount, maxchar) {
    if (inputcheck.value.length > maxchar) { inputcheck.value = inputcheck.value.substring(0, maxchar); }
    else { inputcount.value = maxchar - inputcheck.value.length; }
}
$(".custom-file-input").on("change", function () {
    var fileName = $(this).val().split("\\").pop();
    $(this).siblings(".custom-file-label").addClass("selected").html(fileName);
});
function InsertToText(tid, ivalue) { document.getElementById(tid).value = document.getElementById(tid).value + ivalue; }
function changeValueToIdInput(e, idtarget) { var id = getIdJquery(idtarget); $(id).val($(e).val()); }
function disableEnterKey(e) { var key; if (window.event) { key = window.event.keyCode; } else { key = e.which; } /* firefox */ return (key != 13); }
var isRunLoad = 0;
var path_images = '/images/';
Number.prototype.format = function (n = 0, x = 3) { var re = '\\d(?=(\\d{' + (x || 3) + '})+' + (n > 0 ? '\\.' : '$') + ')'; return this.toFixed(Math.max(0, ~~n)).replace(new RegExp(re, 'g'), '$&,'); };
Number.prototype.formatVN = function (n = 0, x = 3) {
    var re = '\\d(?=(\\d{' + (x || 3) + '})+' + (n > 0 ? '\\.' : '$') + ')';
    re = (this.toFixed(Math.max(0, ~~n)).replace(new RegExp(re, 'g'), '$&,')).replace(/[.]/g, '|').replace(/[,]/g, '.');
    return re.replace(/[|]/g, ',');
};
function formatNumberVNTarget(eInput, eTarget) {
    var tg = $(eInput).parent().find(eTarget);
    var v = $(eInput).val();
    if (v == '') { tg.text(''); return; }
    if (/^[0-9-]+$/g.test(v) == false) { tg.text('N/A'); return; }
    tg.text(parseFloat(v).formatVN());
}
$.fn.insertAt = function (index, $parent) { return this.each(function () { if (index === 0) { $parent.prepend(this); } else { $parent.children().eq(index - 1).after(this); } }); }
function isNumberKey(e) { var charCode = (e.which) ? e.which : e.keyCode; if (charCode != 46 && charCode > 31 && (charCode < 48 || charCode > 57)) return false; return true; }
var keyPressNumber = function (e) { var a = []; var k = e.which; for (i = 48; i < 58; i++) a.push(i); if (!(a.indexOf(k) >= 0)) e.preventDefault(); }
var loadImg = function (event, id) { var output = document.getElementById(id); output.src = URL.createObjectURL(event.target.files[0]); };
function getFileExtension(filename) { var ext = /^.+\.([^.]+)$/.exec(filename); return ext == null ? "" : ext[1]; }
function fixAllClass() {
    $('.clsdate').datepicker({ language: "vi" });
    $('.clstime').timepicker({ 'timeFormat': 'H:i:s' });
}
function vi_en(alias) {
    var str = alias;
    str = str.toLowerCase();
    str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
    str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
    str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
    str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
    str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
    str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
    str = str.replace(/đ/g, "d");
    str = str.replace(/!|@|%|\^|\*|\(|\)|\+|\=|\<|\>|\?|\/|,|\.|\:|\;|\'| |\"|\&|\#|\[|\]|~|$|_/g, "-");
    str = str.replace(/-+-/g, "-");
    str = str.replace(/^\-+|\-+$/g, "");
    return str;
}
function clearUnicode(input) {
    var str = input;
    str = str.toLowerCase();
    str = str.replace(/à|á|ạ|ả|ã|â|ầ|ấ|ậ|ẩ|ẫ|ă|ằ|ắ|ặ|ẳ|ẵ/g, "a");
    str = str.replace(/è|é|ẹ|ẻ|ẽ|ê|ề|ế|ệ|ể|ễ/g, "e");
    str = str.replace(/ì|í|ị|ỉ|ĩ/g, "i");
    str = str.replace(/ò|ó|ọ|ỏ|õ|ô|ồ|ố|ộ|ổ|ỗ|ơ|ờ|ớ|ợ|ở|ỡ/g, "o");
    str = str.replace(/ù|ú|ụ|ủ|ũ|ư|ừ|ứ|ự|ử|ữ/g, "u");
    str = str.replace(/ỳ|ý|ỵ|ỷ|ỹ/g, "y");
    str = str.replace(/đ/g, "d");
    return str;
}
function OnChangeTextNumber(obj) {
    try {
        var t = $(obj).val();
        if (t == '') return;
        var v = parseFloat(t.replace(/,/g, ''));
        $(obj).val(v.format(0, 3));
    }
    catch (error) { $(obj).val('') }
}
var getCookies = function () {
    // Lấy tất cả các cookies hiện có và lưu trữ trong biến cookies
    var cookies = document.cookie.split(";");

    // Tạo một đối tượng lưu trữ các cookie
    var cookieList = {};

    // Duyệt qua tất cả các cookies
    for (var i = 0; i < cookies.length; i++) {
        // Lấy tên và giá trị của cookie
        var cookie = cookies[i].split("=");
        var cookieName = cookie[0].trim();
        var cookieValue = cookie[1];

        // Lưu trữ cookie vào đối tượng cookieList
        cookieList[cookieName] = cookieValue;
    }
    return cookieList;
}
function setActiveLiMenu() {
    var urlCurrent = window.location.href.toLowerCase();
    $('#accordionSidebar').find('li.active').removeClass("active");
    $('#accordionSidebar').find('a.nav-link').each(function () {
        var link = this.href.toLowerCase();
        if (link.endsWith("/")) {
            if (link == urlCurrent) { $(this).parent().addClass("active"); return false; }
        } else {
            if (urlCurrent.startsWith(link)) { $(this).parent().addClass("active"); return false; }
        }
    });
}

function showMessage(sender) {
    if (sender['idmsg'] != undefined) {
        if (sender['idmsg'] != '') {
            idmsg = sender['idmsg'];
            if (!idmsg.startsWith('#')) idmsg = '#' + idmsg;
        }
    }
    var footer = '<a href="javascript:void(0);" class="btn btn-primary btn-sm" data-dismiss="modal"> <i class="fa fa-times"></i> Hủy</a>';
    if (sender['footer'] != undefined) footer = sender['footer'] + " " + footer;
    if (sender['action'] != undefined) footer = '<a href="' + sender['action'] + '" class="btn"> <i class="fa fa-save"></i> Lưu lại </a>' + " " + footer;
    if (sender['title'] != undefined) { $(idmsg).find('.modal-title').html(sender['title']) }
    if (sender['body'] != undefined) { $(idmsg).find('div.modal-body').html(sender['body']) }
    $(idmsg).find('div.modal-footer').html(footer);
    $(idmsg).modal();
    /* $(idmsg).css('display', 'block'); */
}
function showMessageDel(message, url) {
    var body = '<div class="alert alert-warning"> <i class="fa fa-warning"> Bạn thực sự muốn xóa ' + message + ' không? </div>';
    if (message.indexOf('?') > 1) { body = '<div class="alert alert-warning"> <i class="fa fa-warning"> ' + message + ' </div>'; }
    var footer = '<a href="' + url + '" class="btn btn-primary"> <i class="fa fa-check"></i> Có </a>';
    showMessage({ body: body, footer: footer });
}
function showMessageUp(message, url) {
    var body = '<div class="alert alert-info"> Bạn thực sự muốn cập nhật ' + message + ' không? </div>';
    if (message.indexOf('?') > 1) { body = '<div class="alert alert-info"> <i class="fa fa-warning"> ' + message + ' </div>'; }
    var footer = '<a href="' + url + '" class="btn"> <i class="fa fa-check"></i> Có </a>';
    showMessage({ body: body, footer: footer });
}
function messageBox(title, body) { showMessage({ title: title, body: body }); }

function getIdJquery(idObject) {
    if (idObject == undefined) { return ''; }
    if (typeof (idObject) == 'object') { if ($(idObject).length == 0) { return ''; } return idObject; }
    if (typeof (idObject) != 'string') { return ''; }
    if (idObject == '') { return ''; }
    if (/[#]/gi.test(idObject) == false) { idObject = '#' + idObject; }
    return idObject;
}
function getElementJquery(idObject) {
    if (idObject == undefined) { return ''; }
    if (typeof (idObject) == 'object') { if ($(idObject).length == 0) { return ''; } return idObject; }
    if (typeof (idObject) != 'string') { return ''; }
    if (idObject == '') { return ''; }
    if (/[#.]/gi.test(idObject)) { return idObject; }
    if ($('#' + idObject).length > 0) { return '#' + idObject; }
    return '.' + idObject;
}
function postform(fromID, urlPost, targetID, callback) {
    var timestamp = (Math.floor((new Date()).getTime() / 1000)).toString();
    /** Lấy id Element hiển thị */
    var idtarget = "";
    if (typeof (targetID) == 'function') { if (typeof (callback) != 'function') { callback = idtarget; } }
    else if (typeof (targetID) == "object") {
        if ($(targetID).attr("id") == "") { idtarget = "idtarget" + timestamp; $(targetID).attr("id", idtarget); }
    }
    idtarget = getElementJquery(idtarget);

    /** Lấy id From truyền dữ liệu */
    var idform = "";
    if (typeof (fromID) == "object") { if ($(fromID).attr("id") == "") { idform = "idform" + timestamp; $(fromID).attr("id", idform); } }
    var idform = getIdJquery(fromID);

    /** Lấy Url Post */
    var url = "";
    if (typeof (urlPost) == 'string') { url = urlPost; }
    if (url == "") {
        if (idform != "") { if ($(idform).attr("action") != "") { url = $(idform).attr("action"); } }
        if (url == "") { url = window.location.href; }
    }
    if (idform == '') { messageBox('<div class="alert alert-danger">Thông báo lỗi', '<div class="alert alert-danger">Không có nguồn Form để thao tác</div>'); return; }
    if ($(idform).length == 0) { messageBox('<div class="alert alert-danger">Thông báo lỗi', `<div class="alert alert-danger">Không tìm thấy Form có id: ${idform} để thao tác</div>`); return; }
    if ($(idform).attr('enctype') == 'multipart/form-data') {
        var dataform = new FormData(document.getElementById(idform.replace('#', '')));
        if (idtarget != '') { $(idtarget).html('Đang thực hiện <img src="/images/loader.gif" alt="" />'); }
        messageBox('Thông báo', '<div class="alert alert-info">Đang thực hiện <img alt="" title="" src="/images/loader.gif" /></div><progress id="progressBar' + timestamp + '" value="0" max="100" style="width: 100%;"></progress> <span id="progressPercent' + timestamp + '">0%</span>');
        $.ajax({
            url: url, type: "POST", data: dataform,
            mimeTypes: "multipart/form-data",
            contentType: false, cache: false, processData: false
            , xhr: function () {
                const xhr = $.ajaxSettings.xhr();
                if (xhr.upload) {
                    xhr.upload.addEventListener('progress', function (e) {
                        if (e.lengthComputable) {
                            const percentComplete = (e.loaded / e.total) * 100;
                            $('#progressBar' + timestamp).val(percentComplete);
                            $('#progressPercent' + timestamp).text(`${Math.round(percentComplete)}%`);
                        }
                    }, false);
                }
                return xhr;
            }})
            .done(function (response) { ajaxSuccess(response, true, idtarget, callback); })
            .fail(function (jqXHR, textStatus, errorThrown) { ajaxFail(jqXHR, textStatus, errorThrown, idtarget); } );
    }
    else {
        var dataform = $(id).serialize();
        if (idtarget == '') { messageBox('Thông báo', 'Đang thực hiện <img src="/images/loader.gif" alt="" />'); }
        else { $(idtarget).html('Đang thực hiện <img src="/images/loader.gif" alt="" />'); }
        $.ajax({ url: url, type: "POST", data: dataform })
            .done(function (response) { ajaxSuccess(response, true, idtarget, callback); })
            .fail(function (jqXHR, textStatus, errorThrown) { ajaxFail(jqXHR, textStatus, errorThrown, idtarget); });
    }
}
function ajaxSuccess(response, isUpload, idtarget, callback) {
    if (isUpload == false && idtarget == "") { $(idmsg).modal('hide'); }
    if (idtarget != "") { $(idtarget).html(response); }
    if (isUpload) { $(idmsg).find(".modal-body").html(response); }
    else if (idtarget == "") { messageBox('Thông báo', response); }
    if (typeof (callback) == 'function') { callback(); }
    fixAllClass();
}
function ajaxFail(jqXHR, textStatus, errorThrown, idtarget) {
    var tmp = `<div class="alert alert-danger"> JS Lỗi trong quá trình truyền nhận dữ liệu: ${jqXHR.status}: ${textStatus}; ${errorThrown} </div>`;
    if (idtarget == "") { messageBox('<i class="fa fa-warning"></i> JS Thông báo lỗi', tmp); return; }
    else { $(idtarget).html(tmp); }
}
function showgeturl(url, idtarget, callback) {
    if (typeof (idtarget) == 'function') { if (typeof (callback) != 'function') { callback = idtarget; } }
    idtarget = getElementJquery(idtarget);
    var modalshow = false;
    if (typeof (idtarget) == 'string') {
        if (idtarget == '') { messageBox('Thông báo', 'Đang tải dữ liệu <img src="/images/loader.gif" alt="" />'); modalshow = true; }
    }
    if (modalshow == false) { $(idtarget).html('Đang tải dữ liệu <img src="/images/loader.gif" alt="" />'); }
    $.get(url, function (response) {
        if (modalshow) { messageBox('Thông báo', response); }
        else { $(idtarget).html(response); }
        if (typeof (callback) == 'function') { callback(); }
        fixAllClass();
    }).fail(function () {
        if (modalshow) { messageBox('Thông báo', 'Lỗi trong quá trình truyền nhận dữ liệu'); return; }
        $(idtarget).html('Lỗi trong quá trình truyền nhận dữ liệu');
    });
}
function showForm(urlinfo, idform) {
    var idmsg = '#modal-message';
    var title = "Thông tin";
    var body = "";
    var footer = '<a href="javascript:void(0);" class="btn btn-primary btn-sm" data-dismiss="modal"> <i class="fa fa-times"></i> Hủy</a>';
    footer = '<a href="javascript:submitidform(\'' + idform + '\');" class="btn"> <i class="fa fa-save"></i> Cập nhập </a>' + " " + footer;
    $(idmsg).find('div.modal-body').html('<img alt="" src="/images/loader.gif"/>');
    $(idmsg).find('div.modal-body').load(urlinfo);
    $(idmsg).find('h4.modal-title').first().text(title);
    $(idmsg).find('div.modal-footer').html(footer);
    $(idmsg).modal();
}
function GetInfomation(url, eload) {
    eload = eload || '';
    $.getJSON(url, function (data) {
        if (data != null) { for (var key in data) { var id = '#' + key; $(id).val(data[key]); } }
        if (eload == '') { return; }
        if (!eload.startsWith('#') || !eload.startsWith('.')) eload = '#' + eload;
        $(eload).remove();
    });
}
function selectTextToClass(e, nameClass) {
    if (typeof (nameClass) != 'string') { return; }
    var v = $(e).find('option:selected').first().text();
    if (/^[.]/.test(nameClass) == false) { nameClass = "." + nameClass; }
    $(nameClass).text(v);
}
function setsuggest(idInput, tab, field) {
    if (!idInput.startsWith('#')) { idInput = '#' + idInput; }; if (typeof (field) != 'string') { field = ''; }
    $(idInput).autocomplete({
        source: function (request, response) {
            $.ajax({
                url: "/ajax/suggest.php", method: "POST", dataType: "json",
                data: { to: tab, key: $(idInput).val(), f: field }, success: function (data) { response(data); }
            });
        },
        minLength: 2,
        maxShowItems: 6
    });
}
function suggest(idObject, tab, field) {
    if (typeof (idObject) == 'string') { if (idObject == '') { return; } idObject = getIdJquery(idObject); }
    if (typeof (field) != 'string') { field = ''; } if (field == '') { return; }
    $(idObject).autocomplete({
        source: function (request, response) {
            $.ajax({
                url: "/ajax/suggest2.php", method: "POST", dataType: "json",
                data: { t: tab, f: field, v: $(idObject).val() }, success: function (data) { response(data); }
            });
        },
        minLength: 2,
        maxShowItems: 6
    });
}
function callLogout() { showgeturl("/login/logout"); }
function checkAllToName(e, targetName) { $('[name^="' + targetName + '"]').prop('checked', $(e).prop('checked')); }
function autocomplete(sender, arr) {
    var typesender = typeof (sender);
    if (typesender == 'string') {
        var idobject = getElementJquery(sender);
        $(idobject).each(function () { autocomplete(this, arr); });
        return;
    }
    var currentFocus;
    sender.addEventListener("input", function (e) {
        var a, b, i, val = this.value;
        closeAllLists();
        if (!val) { return false; }
        currentFocus = -1;
        a = document.createElement("DIV");
        a.setAttribute("id", this.id + "autocomplete-list");
        a.setAttribute("class", "autocomplete-items");
        this.parentNode.appendChild(a);
        for (i = 0; i < arr.length; i++) {
            if (arr[i].substr(0, val.length).toUpperCase() == val.toUpperCase()) {
                b = document.createElement("DIV");
                b.innerHTML = "<strong>" + arr[i].substr(0, val.length) + "</strong>";
                b.innerHTML += arr[i].substr(val.length);
                b.innerHTML += "<input type='hidden' value='" + arr[i] + "'>";
                b.addEventListener("click", function (e) {
                    sender.value = this.getElementsByTagName("input")[0].value;
                    closeAllLists();
                });
                a.appendChild(b);
            }
        }
    });
    sender.addEventListener("keydown", function (e) {
        var x = document.getElementById(this.id + "autocomplete-list");
        if (x) x = x.getElementsByTagName("div");
        if (e.keyCode == 40) { currentFocus++; addActive(x); }
        else if (e.keyCode == 38) { currentFocus--; addActive(x); }
        else if (e.keyCode == 13) { e.preventDefault(); if (currentFocus > -1) { if (x) x[currentFocus].click(); } }
        else if (e.keyCode == 9) { if (currentFocus > -1) { if (x) x[currentFocus].click(); } closeAllLists(); }
    });
    function addActive(x) {
        if (!x) return false;
        removeActive(x);
        if (currentFocus >= x.length) currentFocus = 0;
        if (currentFocus < 0) currentFocus = (x.length - 1);
        x[currentFocus].classList.add("autocomplete-active");
    }
    function removeActive(x) {
        for (var i = 0; i < x.length; i++) { x[i].classList.remove("autocomplete-active"); }
    }
    function closeAllLists(elmnt) {
        var x = document.getElementsByClassName("autocomplete-items");
        for (var i = 0; i < x.length; i++) {
            if (elmnt != x[i] && elmnt != sender) { x[i].parentNode.removeChild(x[i]); }
        }
    }
    document.addEventListener("click", function (e) { closeAllLists(e.target); });
}