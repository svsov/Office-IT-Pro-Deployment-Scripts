$(document).ready(function() {


    $("#btViewOnGitHub").click(function () {
        window.location.href = "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts";
        return false;
    });

    $("#btDownloadZip").click(function () {
        window.location.href = "https://github.com/OfficeDev/Office-IT-Pro-Deployment-Scripts/zipball/master";
        return false;
    });


    resizeWindow();
});

function resizeWindow() {
    var bodyHeight = window.innerHeight;
    var bodyWidth = window.innerWidth;
    var leftPaneHeight = bodyHeight - 180;

    var iframeHeight = bodyHeight - 80;
    $("#mainFrame").height(iframeHeight);

}
