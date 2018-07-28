import QtQuick 2.4
import QtQuick.XmlListModel 2.0
import QtQml.Models 2.2

Rectangle {
    id: root
    //width: 300; height: 400
    height:  imag.height
    width:  imag.width
    property int index: 0
   // color: "black"
    Image {
        id: imag
        source: "qrc:/1.jpg"
    }
    Timer
    {
        id:timer
        interval: 1000;running: true; repeat: true
        onTriggered: {
            console.log("hi")
            if( 1 === index )
            imag.source = "qrc:/1.jpg"


            if( 2 === index )
            imag.source = "qrc:/2.jpg"

            if( 3 === index )
            {
                imag.source = "qrc:/3.jpg"
                index = 0;
            }
            index++;
        }
    }

    Component.onCompleted:
    {
        console.log("loaded")
        timer.start()
    }
}
