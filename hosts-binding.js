/*********************************************************************\
 Filename: hosts-binding.js
 Author:   Wang Zhenwu (zhenwu.wang@gmail.com)
 Version:  1.1.0
 Purpose:  manipulate hosts binding in /etc/hosts files.
 Usage:    Just "double click", It will prompt for userright if
           UAC control is enabled.

 ChangeLog:
 + 2014-07-01 Wang Zhenwu(zhenwu.wang@gmail.com) 1.1.0
 - Feature: use folder for backuped hosts file.

 + 2014-07-01 Wang Zhenwu(zhenwu.wang@gmail.com) 1.0.1
 - Bugfix: when backup(move) the hosts file, the file owner is
           changed, so hosts file not work anymore. now, use copy to
           backup the hosts file, then overwrite the original file.

 + 2014-06-28 Wang Zhenwu(zhenwu.wang@gmail.com) 1.0.0
 - Feature: use table view, so user can binding/unbinding one host

 + 2010-07-01 Wang Zhenwu(zhenwu.wang@gmail.com) 0.1.0
 - Feature: create a script comment the hosts binding.

\*********************************************************************/

var ready = false;
var test = 'undefined';
var reBlank = /^\s*$/;
var changeFile = false;

var ie;
var bindHosts = {};
var hostsLines = [];
var l_num = 0;


if (WScript.Arguments.Count() == 0) {
   var oShell = new ActiveXObject("Shell.Application");
/*    oShell.ShellExecute("wscript.exe",
                        "\"" + WScript.ScriptFullName + "\"" +
                            " /isElevated",
                        WScript.ScriptFullName.slice(0,
                            -WScript.ScriptName.length-1),
                        "runas", 1);
*/
    oShell.ShellExecute("wscript.exe",
                        "\"" +  WScript.ScriptFullName +  "\"" +
                                " RunAsAdministrator",
                        "",
                        "runas", 1);

} else {
    var ws = WScript.CreateObject("WScript.Shell");
    var SysRoot = ws.ExpandEnvironmentStrings("%SystemRoot%");
    var HomePath = ws.ExpandEnvironmentStrings("%HOMEPATH%");
    var UserName = ws.ExpandEnvironmentStrings("%USERNAME%");
    var TempPath = ws.ExpandEnvironmentStrings("%TEMP%");

    var fso, f1, f2, f3;
    var ForReading = 1;
    var ForWriting = 2;
    var iHTML = "";

    fso = new ActiveXObject("Scripting.FileSystemObject");

    var backupDir = HomePath + "\\Desktop\\" + UserName + ".Hosts.bak";

    if (!fso.FolderExists(backupDir)) {
        fso.CreateFolder(backupDir);
    }

    var hostsFileOrig = SysRoot + "\\System32\\drivers\\etc\\hosts";
    var hostsFileBack = backupDir + "\\" + "hosts." + 
                        getTimestamp() + ".txt";

    var re = /^\s*(#*)\s*([0-9]+\.[0-9]+\.[0-9]+\.[0-9]+)\s+([^#]+)\b\s*#?\s*(.*)\s*$/;

    f1 = fso.OpenTextFile(hostsFileOrig, ForReading, true);
    //f3 = fso.CreateTextFile(hostsFileNew, true);

    iHTML += '<!doctype html><html>';
    iHTML += '<head><title>hosts</title>';
    iHTML += '<meta charset="utf-8"/>';
    iHTML += '<style type="text/css">';
    iHTML += 'td.s {background: #8CB2E1;}';
    iHTML += 'td.scenter {text-align: center; background: #8CB2E1;}';
    iHTML += 'td.d {background: #90D2E7;}';
    iHTML += 'td.dcenter {text-align: center;background: #90D2E7;}';
    iHTML += 'td.center {text-align: center;}';
    iHTML += 'th {background: #8F8ADE;}';
    iHTML += '</style>';
    iHTML += '</head>';
    iHTML += '<body><form name="hostform">';
    iHTML += '<table>';
    iHTML += '<tr><th>��</th><th>ȡ��</th>';
    iHTML += '<th>ɾ��</th><th>IP��ַ</th>';
    iHTML += '<th>������</th><th>˵��</th></tr>';

    l_num = 0;
    var counter = 0;
    while (! f1.AtEndOfStream) {
        s = f1.ReadLine();
        if (re.exec(s)) {
            Commented = RegExp.$1;
            IPAddr = RegExp.$2;
            Host = RegExp.$3;
            Desc = RegExp.$4;
            var tdClass;
            if (counter % 2 == 0) {
                tdClass = 'd';
            } else {
                tdClass = 's';
            }
            var checkedCancel = "";
            var checkedBind = "";
            if (reBlank.exec(Commented)) {
                checkedBind = 'checked="checked"';
            } else {
                checkedCancel = 'checked="checked"';
            }
            iHTML += '<tr><td class="' + tdClass + 'center">';
            iHTML += '<input type="radio" name="l' + l_num +
                     '" value="l' + l_num + '_bind" ' + checkedBind +
                     ' /></td>';
            iHTML += '<td class="'+ tdClass + 'center">';
            iHTML += '<input type="radio" name="l' + l_num +
                     '" value="l' + l_num + '_cancel" ' + 
                     checkedCancel +
                     ' /></td>';
            iHTML += '<td class="' + tdClass + 'center">';
            iHTML += '<input type="radio" name="l' + l_num +
                         '" value="l' + l_num + '_delete" /></td>';
            iHTML += '<td class="' + tdClass +'">' + IPAddr + '</td>';
            iHTML += '<td class="' + tdClass +'">' + Host + '</td>';
            iHTML += '<td class="' + tdClass +'">' + 
                     Desc + '</td></tr>' ;
            counter ++;
            if (reBlank.exec(Desc)) {
                hostsLines[l_num] = IPAddr + "\t" + Host;
            } else {
                hostsLines[l_num] = IPAddr + "\t" + Host + ' #' + Desc;
            }
            bindHosts[l_num] = 0;
        } else {
            hostsLines[l_num] = s;
        }
        l_num ++;
    }

    f1.close();

    iHTML += '<tr><td colspan="6">&nbsp;</td></tr>';
    iHTML += '<tr><td colspan="6" class="center">';
    iHTML += '<input type="button" ';
    iHTML += 'name="butn_ok" value="ȷ���ύ" />';
    iHTML += '&nbsp;&nbsp;&nbsp;';
    iHTML += '<input type="button" ';
    iHTML += 'name="butn_bind" value="ȫ����" />';
    iHTML += '&nbsp;&nbsp;&nbsp;';
    iHTML += '<input type="button" ';
    iHTML += 'name="butn_cancel" value="ȫ��ȡ��" />';
    iHTML += '&nbsp;&nbsp;&nbsp;';
    iHTML += '<input type="button" ';
    iHTML += 'name="butn_delete" value="ȫ��ɾ��" />';
    iHTML += '</td></tr>';

    iHTML += '<tr><td colspan="6">';
    iHTML += '<textarea name="newbind" rows="10" cols="80">';
    iHTML += '</textarea></td></tr>';
    iHTML += '<tr><td colspan="6" class="center">';
    iHTML += '<input type="button" ';
    iHTML += 'name="butn_add" value="������" />';
    iHTML += '&nbsp;&nbsp;&nbsp;';
    iHTML += '<input type="button" ';
    iHTML += 'name="butn_close" value="�رմ���" />';
    iHTML += '</td></tr>';
    iHTML += '</table></form></body></html>';

    ie = WScript.CreateObject("InternetExplorer.Application", "IE_");
    ie.Navigate("about:blank");
    ie.Visible = false;
    ie.MenuBar = false;
    ie.ToolBar = false;
    ie.StatusBar = false;
    ie.document.write(iHTML);

    if (ie.Busy) {
        WScript.Sleep(1000);
    }

    ready = false;
    ie.document.hostform.butn_ok.onmouseup = butn_ok;
    ie.document.hostform.butn_bind.onmouseup = butn_bind;
    ie.document.hostform.butn_cancel.onmouseup = butn_cancel;
    ie.document.hostform.butn_delete.onmouseup = butn_delete;
    ie.document.hostform.butn_add.onmouseup = butn_add;
    ie.document.hostform.butn_close.onmouseup = butn_close;
    ie.Visible = true;

    while (!ready) {
        WScript.Sleep(500);
    }

    if (changeFile) {
        f1 = fso.GetFile(hostsFileOrig);
        f1.Copy(hostsFileBack);

        f2 = fso.CreateTextFile(hostsFileOrig, true);
        for (var l1 in hostsLines) {
            if (bindHosts.hasOwnProperty(l1)){
                if (bindHosts[l1] == 1) {
                    f2.WriteLine(hostsLines[l1]);
                } else if (bindHosts[l1] == 2) {
                    f2.WriteLine('#' + hostsLines[l1]);
                } else if (bindHosts[l1] == 3) {
                    continue;
                } else {
                    f2.WriteLine(hostsLines[l1]);
                }
            } else {
                f2.WriteLine(hostsLines[l1]);
            }
        }
        f2.close();
    }
}

function IE_OnQuit()
{
    ready = true;
}


function getTimestamp()
{
    var curDate = new Date();

    var year = curDate.getFullYear();
    var month = curDate.getMonth() + 1;
    var day = curDate.getDate();
    var hour = curDate.getHours();
    var min = curDate.getMinutes();
    var sec = curDate.getSeconds();
    var ms = curDate.getMilliseconds();
    if (month < 10) month = "0" + month;
    if (day < 10) day = "0" + day;
    if (hour < 10) hour = "0" + hour;
    if (min < 10) min = "0" + min;
    if (sec < 10) sec = "0" + sec;
    if (ms < 10) {
        ms = "00" + ms;
    } else if (ms < 100) {
        ms = "0" + ms;
    }
    return "" + year + month + day + hour + min + sec + '.' + ms;
}


function butn_close()
{
    ready = true;
    ie.Quit();
}


function butn_ok()
{
    ready = true;
    changeFile = true;
    ie.Quit();
    for (var k1 in bindHosts) {
        if (ie.document.getElementsByName('l' + k1)[0].checked) {
            bindHosts[k1] = 1;
        } else if (ie.document.getElementsByName('l' + k1)[1].checked) {
            bindHosts[k1] = 2;
        } else if (ie.document.getElementsByName('l' + k1)[2].checked) {
            bindHosts[k1] = 3;
        }
    }
}


function butn_bind()
{
    ready = true;
    changeFile = true;
    ie.Quit();
    for (var k1 in bindHosts) {
        bindHosts[k1] = 1;
    }
}



function butn_cancel()
{
    ready = true;
    changeFile = true;
    ie.Quit();
    for (var k1 in bindHosts) {
        bindHosts[k1] = 2;
    }
}


function butn_delete()
{
    ready = true;
    changeFile = true;
    ie.Quit();
    for (var k1 in bindHosts) {
        bindHosts[k1] = 3;
    }
}


function butn_add()
{
    ready = true;
    changeFile = true;
    ie.Quit();

    for (var k1 in bindHosts) {
        if (ie.document.getElementsByName('l' + k1)[0].checked) {
            bindHosts[k1] = 1;
        } else if (ie.document.getElementsByName('l' + k1)[1].checked) {
            bindHosts[k1] = 2;
        } else if (ie.document.getElementsByName('l' + k1)[2].checked) {
            bindHosts[k1] = 3;
        }
    }

    var newBinds = ie.document.getElementsByName("newbind")[0].value;
    var newHosts = newBinds.split("\n");

    for (var h1 in newHosts) {
        if (reBlank.exec(newHosts[h1])) {
            continue;
        } else {
            hostsLines[l_num] = newHosts[h1];
            l_num ++;
        }
    }
}
