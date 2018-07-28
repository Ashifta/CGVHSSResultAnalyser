import QtQuick 2.6
import QtQuick.Window 2.2
import QtQuick.Controls 2.0
import QtQuick 2.2
import QtQuick.Dialogs 1.0
import "./"

Window {
    visible: true
    id:root

    maximumWidth: 1000
    maximumHeight: 475
    width: 1000
    height: 475
//    maximumWidth: 1100
//    maximumHeight:  700

    minimumWidth:  1100
    minimumHeight: 475
    title: qsTr("CGVHSS Result Analyser Ver 1.2 ")
    color: "#1D4580"
    MouseArea {
        anchors.fill: parent
        onClicked: {
            //console.log(qsTr('Clicked on background. Text: "' + textEdit.text + '"'))
        }
    }

    FileDialog {
        id: fileDialog
        title: "Please choose a file"
        nameFilters: ["Excel file(*.xlsx)", "Excel file(*.xls)"]
        folder: shortcuts.home
        onAccepted: {
            console.log("You chose: " + fileDialog.fileUrl)
            filePath.text = fileDialog.fileUrl
            schoolModel.setFilePath(fileDialog.fileUrl, modelCombo.currentIndex)
        }
        onRejected: {
            console.log("Canceled")
            Qt.quit()
        }
        Component.onCompleted: visible = false
    }


    Rectangle
    {
        width: parent.width
        height: parent.height
        color: "#EBEDEC"

    }


    Column
    {
        anchors.verticalCenter:parent
        Row
        {

            anchors.horizontalCenter: parent.horizontalCenter
            Image {
                id: name

                source: "qrc:/2017-11-16.png"
                 width: 50
                 height:50
             //   opacity: .2
            }

            Text
            {
                color: "black"
                font.pixelSize: 20
                text:"Calicut Girls' Vocational & Higher Secondary School\nKundungal, Kallai P.O, 673003, www.calicutgirlsschool.org"
            }
        }
        Rectangle
        {
            width: root.width
            height: 5
            color:"#393187"
        }
        Text
        {
            anchors.horizontalCenter:parent.horizontalCenter
            color: "black"
            font.pixelSize: 20
            text:"All in One Result Analyser 1.0(SSLC, HSS, VHSS)"
        }

        Rectangle
        {
            width: root.width
            height: 5
            color:"#393187"
        }


        Row
        {
           // anchors.horizontalCenter: parent.horizontalCenter
             spacing: 5
                    Text
                    {
                     //   anchors.horizontalCenter: parent.horizontalCenter
                        width: root.width/2
                        id:textHeight
                        color: "black"
                        font.pixelSize: 12
                        text:"
    Steps to Follow: \n
    1.  Open the link http://keralaresults.nic.in or similar sites.
    2.  Give the school code, and Press Enter, the result will be displayed.
    3.  Copy the marks from the first roll number (including the Register Number)
         onwards to wherever you required to analyse.
    4.  Open the MS Office Excel & Paste the Content (Ctrl V).(Copy to the top first line).
    5.  Create a second Work Sheet in the Excel file by clicking at the lower left
         side of the file.
    6.  Save it with required Name and CLOSE the Excel file.
    7.  Select Section (SSLC, HSS-PlusOne, HSS-PlusTwo,VHSS-FirstYear, VHSS-SecondYear)
         from the All in One Result Analyser.
    8.  Brows the file you need to analyse.
    9.  Press GenerateReport Button.
    10. After displaying 'Result Generated', open the same file that you saved.
    11. Result will be available on 2nd Work Sheet.\n\n\n
        **If you find the \"CGVHSS Result Analyser\* is not responding, Kindly restart the Application"


                    }

                    Rectangle
                    {
                        color: "#393187"
                        width: 5
                        height: parent.height
                    }

                    Column
                    {
                        anchors.verticalCenter:  parent.verticalCenter

                        Text {
                            id: name4
                            text: qsTr("\nFOLLOWING RESULTS WILL BE GENERATED\n
        1. Student RankList based on grade points
        2. Total students
        3. Total Pass
        4. Total Full A+
        5. Total Fail
        6. Pass Percentage")
                        }
                        Rectangle
                    {
                    //    anchors.verticalCenter:  parent.verticalCenter
                       // anchors.horizontalCenter: parent.horizontalCenter

                        color: "#393187"
                        width: colum.width+20
                        height: colum.height+20
                    Column
                    {
                        spacing: 5
                        id:colum
                        anchors.verticalCenter:  parent.verticalCenter
                        anchors.horizontalCenter: parent.horizontalCenter
                        ComboBox {
                            width:root.width/6
                            id:modelCombo
                            model: ["SSLC", "HSS-PlusOne", "HSS-PlusTwo","VHSS-FirstYear", "VHSS-SecondYear"]


                        }
                        Row
                        {
                        TextField
                        {
                            width:root.width/4
                            id:filePath
                            placeholderText: "Path to Excel file"
                        }
                        Button
                        {
                            text:"Browse"
                            onClicked:
                            {
                                fileDialog.open()
                            }
                        }
                        }
                        Button
                        {
                            //anchors.left: middle.left
                            text: "Generate Report"
                            onClicked:
                            {
                                if( filePath.text === "")
                                {
                                    return;
                                }

                                progress.text = "Report Generation is in Progress, It may takes 1 to 2 minutes..."
                                progress.color = "green"
                                time.starting  = true
                                progressBar.visible = true;
                                time.start();
                            }
                        }

                    }



                    }
                        Column
                        {
                        Text
                        {
                            //width:root.width/2
                            id:progress
                            color: "black"
                            font.pixelSize: 12
                            font.bold: true
                        }
                        ProgressBar
                        {
                            visible: false
                            indeterminate: true
                            width: root.width/4
                            height: 10
                            id:progressBar
                        }
                        }

                    Text {
                        id: name1
                         font.pixelSize: 12
                        text: qsTr("Note:
If you want to generate a Rank List for a specific course (Eg. Humanities),
copy the students’ details of the Humanities students only, and then open a
fresh Excel file, paste it, save it. Add a Second Sheet.  Select this file for
generating Rank List.
")
                    }
                    }


                    Timer
                    {
                        property bool starting: false
                        id:time
                        interval: 500; running: false; repeat: false
                        onTriggered:
                        {
                            schoolModel.generateReport(modelCombo.currentIndex)
                        }
                    }


        }
        Rectangle
        {
            width: root.width
            height: 5
            color:"#393187"
        }
        Column
        {
            anchors.horizontalCenter:parent.horizontalCenter
            Text
            {
                anchors.horizontalCenter:parent.horizontalCenter
                //width:root.width/2
                id:copyright
                color: "black"
                text: "Sofware Developed by Ashif T.A ( Contact : ashiftaec@gmail.com )"

            }

        }
    }


    Connections
    {
        target: schoolModel

        onSheetIsNotAvalable:{
            console.log("temp")
            progress.text = "Please Add sheet number 2 in excel, and run again"
            progress.color = "red"
            progressBar.visible = false;

        }
        onProgress:{
            if(isProgress)
            {
                progress.text = "Successfully generated report"
                progress.color = "green"
                progressBar.visible = false;
            }
        }
    }
}
