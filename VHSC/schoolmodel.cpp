#include "schoolmodel.h"
#include <QtDebug>
#include <QAxObject>
#include <QDir>
#include <QThread>

SchoolModel::SchoolModel()
{
    m_gradeMap.insert("A+", 9 );
    m_gradeMap.insert("A", 8 );

    m_gradeMap.insert("B+", 7 );
    m_gradeMap.insert("B", 6 );

    m_gradeMap.insert("C+", 5 );
    m_gradeMap.insert("C", 4 );

    m_gradeMap.insert("D+", 3 );
    m_gradeMap.insert("D", 2 );

    m_gradeMap.insert("E+", 1 );
    m_gradeMap.insert("E", 0 );

}

void SchoolModel::setFilePath(QString filePath)
{
    m_filePath = filePath;

    qDebug() << m_filePath << m_scoolType;
}

void SchoolModel::generateReport(int type)
{
    m_APlus10Count=0;
    m_APlus9Count=0;
    m_Pass = 0;
    m_Fail = 0;
    m_rankInfo.clear();


    emit progress(true);
    m_scoolType = type;
    switch (m_scoolType) {
    case 0:
        highSchool();
        break;
    case 1:
        hss1();
        break;
    case 2:
        hss();
        break;
    case 3:
        vhsc1();
        break;
    case 4:
        vhsc();
        break;
    default:
        break;
    }
}
////0 = vhsc1 3
////1 = vhsc 4
////2 = hss    1
//// 3 highschool/sslc   0
void SchoolModel::reportFullAPlus()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );

    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* workbook = workbooks->querySubObject( "Open(const QString&)", m_filePath );
    QAxObject* sheets = workbook->querySubObject( "Worksheets" );

    //worksheets count
    int count = sheets->dynamicCall("Count()").toInt();
    sheets->dynamicCall("EditDirectlyInCell(const int)", true);

    count = sheets->property("Count").toInt();
    qDebug()<< "count" << count;
    if(count <= 1)
    {
        emit sheetIsNotAvalable();
        return;
    }

    for (int i=2; i<=2; i++) //cycle through sheets
    {
        //sheet pointer
        QAxObject* sheet = sheets->querySubObject( "Item( int )", i );
        qDebug()<< "sheets2" << sheets;

        QAxObject* rows = sheet->querySubObject( "Rows" );

        int rowCount = rows->dynamicCall( "Count()" ).toInt(); //unfortunately, always returns 255, so you have to check somehow validity of cell values

        QAxObject *cellA, *cellB,*cellC, *cellD, *cellE,*cellF, *cellG, *cellH, *cellI, *cellJ;

        int row =3;

        QString RANK ="A"+QString::number(1);
        cellA = sheet->querySubObject("Range(QVariant, QVariant)",RANK);
        cellA->dynamicCall("SetValue(const QVariant&)",QVariant("Rank List Based to Grade Points"));


        QString A ="A"+QString::number(2);
        cellA = sheet->querySubObject("Range(QVariant, QVariant)",A);
        cellA->dynamicCall("SetValue(const QVariant&)",QVariant("Register No"));

        QString B ="B"+QString::number(2);
        cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
        cellB->dynamicCall("SetValue(const QVariant&)",QVariant("Name"));

        if( m_scoolType != 1 && m_scoolType != 3 )
        {
        QString C ="C"+QString::number(2);
        cellC = sheet->querySubObject("Range(QVariant, QVariant)",C);
        cellC->dynamicCall("SetValue(const QVariant&)",QVariant("A+"));

        QString D ="D"+QString::number(2);
        cellC = sheet->querySubObject("Range(QVariant, QVariant)",D);
        cellC->dynamicCall("SetValue(const QVariant&)",QVariant("A"));

        QString E ="E"+QString::number(2);
        cellC = sheet->querySubObject("Range(QVariant, QVariant)",E);
        cellC->dynamicCall("SetValue(const QVariant&)",QVariant("B+"));

        QString F ="F"+QString::number(2);
        cellC = sheet->querySubObject("Range(QVariant, QVariant)",F);
        cellC->dynamicCall("SetValue(const QVariant&)",QVariant("B"));

        QString G ="G"+QString::number(2);
        cellC = sheet->querySubObject("Range(QVariant, QVariant)",G);
        cellC->dynamicCall("SetValue(const QVariant&)",QVariant("C+"));

        QString H ="H"+QString::number(2);
        cellA = sheet->querySubObject("Range(QVariant, QVariant)",H);
        cellA->dynamicCall("SetValue(const QVariant&)",QVariant("C"));

        if( m_scoolType == 0 )
        {
            QString I ="I"+QString::number(2);
            cellA = sheet->querySubObject("Range(QVariant, QVariant)",I);
            cellA->dynamicCall("SetValue(const QVariant&)",QVariant("D+"));

            QString J ="J"+QString::number(2);
            cellA = sheet->querySubObject("Range(QVariant, QVariant)",J);
            cellA->dynamicCall("SetValue(const QVariant&)",QVariant("D"));
        }
        else if( m_scoolType ==2 )
        {
            QString I ="I"+QString::number(2);
            cellA = sheet->querySubObject("Range(QVariant, QVariant)",I);
            cellA->dynamicCall("SetValue(const QVariant&)",QVariant("Total Marks"));
        }

        }
        else
        {
            QString C ="C"+QString::number(2);
            cellC = sheet->querySubObject("Range(QVariant, QVariant)",C);
            cellC->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 1"));

            QString D ="D"+QString::number(2);
            cellC = sheet->querySubObject("Range(QVariant, QVariant)",D);
            cellC->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 2"));

            QString E ="E"+QString::number(2);
            cellC = sheet->querySubObject("Range(QVariant, QVariant)",E);
            cellC->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 3"));

            if( m_scoolType == 1 )
            {
            QString F ="F"+QString::number(2);
            cellC = sheet->querySubObject("Range(QVariant, QVariant)",F);
            cellC->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 4"));

            QString G ="G"+QString::number(2);
            cellC = sheet->querySubObject("Range(QVariant, QVariant)",G);
            cellC->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 5"));

            QString H ="H"+QString::number(2);
            cellA = sheet->querySubObject("Range(QVariant, QVariant)",H);
            cellA->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 6"));

            }
            else
            {
                QString F ="F"+QString::number(2);
                cellC = sheet->querySubObject("Range(QVariant, QVariant)",F);
                cellC->dynamicCall("SetValue(const QVariant&)",QVariant("Module 1"));

                QString G ="G"+QString::number(2);
                cellC = sheet->querySubObject("Range(QVariant, QVariant)",G);
                cellC->dynamicCall("SetValue(const QVariant&)",QVariant("Module 2"));

                QString H ="H"+QString::number(2);
                cellC = sheet->querySubObject("Range(QVariant, QVariant)",H);
                cellC->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 4"));

                QString I ="I"+QString::number(2);
                cellC = sheet->querySubObject("Range(QVariant, QVariant)",I);
                cellC->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 5"));

                QString J ="J"+QString::number(2);
                cellA = sheet->querySubObject("Range(QVariant, QVariant)",J);
                cellA->dynamicCall("SetValue(const QVariant&)",QVariant("SUBJECT 6"));

            }



            if( m_scoolType == 3 )
            {

                QString K ="K"+QString::number(2);
                cellA = sheet->querySubObject("Range(QVariant, QVariant)",K);
                cellA->dynamicCall("SetValue(const QVariant&)",QVariant("Total Marks"));
            }
            else
            {
                QString I ="I"+QString::number(2);
                cellA = sheet->querySubObject("Range(QVariant, QVariant)",I);
                cellA->dynamicCall("SetValue(const QVariant&)",QVariant("Total Marks"));
            }



        }
        QMultiMap<int, QMultiMap<unsigned long long, RankInfo>>::const_iterator k = m_rankInfo.constEnd();
        while (k != m_rankInfo.constBegin()) {
            --k;

            QMultiMap<unsigned long long, RankInfo>::const_iterator j = k.value().constEnd();
            while (j != k.value().constBegin()) {
                --j;



                QString A ="A"+QString::number(row);
                cellA = sheet->querySubObject("Range(QVariant, QVariant)",A);
                cellA->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().rollNumber));


                QString B ="B"+QString::number(row);
                cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
                cellB->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().name));

                if( m_scoolType == 1 || m_scoolType == 3 )
                {
                    QString C ="C"+QString::number(row);
                    cellC = sheet->querySubObject("Range(QVariant, QVariant)",C);
                    cellC->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["A+"]));

                    QString D ="D"+QString::number(row);
                    cellD = sheet->querySubObject("Range(QVariant, QVariant)",D);
                    cellD->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["A"]));

                    QString E ="E"+QString::number(row);
                    cellE = sheet->querySubObject("Range(QVariant, QVariant)",E);
                    cellE->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["B+"]));

                    QString F ="F"+QString::number(row);
                    cellF = sheet->querySubObject("Range(QVariant, QVariant)",F);
                    cellF->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["B"]));

                    QString G ="G"+QString::number(row);
                    cellG = sheet->querySubObject("Range(QVariant, QVariant)",G);
                    cellG->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["C+"]));

                    QString H ="H"+QString::number(row);
                    cellH = sheet->querySubObject("Range(QVariant, QVariant)",H);
                    cellH->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["C"]));

                    if( m_scoolType == 3 )
                    {
                        QString I ="I"+QString::number(row);
                        cellG = sheet->querySubObject("Range(QVariant, QVariant)",I);
                        cellG->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["D+"]));

                        QString J ="J"+QString::number(row);
                        cellH = sheet->querySubObject("Range(QVariant, QVariant)",J);
                        cellH->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["D"]));


                        QString K ="K"+QString::number(row);
                        cellC = sheet->querySubObject("Range(QVariant, QVariant)",K);
                        cellC->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().mark));
                    }
                    else
                    {
                        QString I ="I"+QString::number(row);
                        cellC = sheet->querySubObject("Range(QVariant, QVariant)",I);
                        cellC->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().mark));
                    }
                    row++;
                    continue;
                }
                QString C ="C"+QString::number(row);
                cellC = sheet->querySubObject("Range(QVariant, QVariant)",C);
                cellC->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["A+"]));

                QString D ="D"+QString::number(row);
                cellD = sheet->querySubObject("Range(QVariant, QVariant)",D);
                cellD->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["A"]));

                QString E ="E"+QString::number(row);
                cellE = sheet->querySubObject("Range(QVariant, QVariant)",E);
                cellE->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["B+"]));

                QString F ="F"+QString::number(row);
                cellF = sheet->querySubObject("Range(QVariant, QVariant)",F);
                cellF->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["B"]));

                QString G ="G"+QString::number(row);
                cellG = sheet->querySubObject("Range(QVariant, QVariant)",G);
                cellG->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["C+"]));

                QString H ="H"+QString::number(row);
                cellH = sheet->querySubObject("Range(QVariant, QVariant)",H);
                cellH->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["C"]));
                if( m_scoolType == 0 )//High School
                {
                    QString I ="I"+QString::number(row);
                    cellG = sheet->querySubObject("Range(QVariant, QVariant)",I);
                    cellG->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["D+"]));

                    QString J ="J"+QString::number(row);
                    cellH = sheet->querySubObject("Range(QVariant, QVariant)",J);
                    cellH->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().gradeMap["D"]));
                }

                if( m_scoolType == 2 )//HSS
                {
                    QString I ="I"+QString::number(row);
                    cellG = sheet->querySubObject("Range(QVariant, QVariant)",I);
                    cellG->dynamicCall("SetValue(const QVariant&)",QVariant(j.value().mark));
                }


                row++;
            }


            QString L, M, str;
            QAxObject * cellL, *cellM;
            if( m_scoolType == 1 || m_scoolType == 3 )
            {

            }
            else
            {
                L ="L"+QString::number(3);
                cellL = sheet->querySubObject("Range(QVariant, QVariant)",L);
                cellL->dynamicCall("SetValue(const QVariant&)",QVariant("Total Full A+"));

                M ="M"+QString::number(3);
                cellM = sheet->querySubObject("Range(QVariant, QVariant)",M);
                cellM->dynamicCall("SetValue(const QVariant&)",QVariant(m_APlus10Count));

                L ="L"+QString::number(4);
                cellL = sheet->querySubObject("Range(QVariant, QVariant)",L);

                if( m_scoolType == 3 )//VHSC1
                {
                    str.append("Total 7 A+");
                }
                else if( m_scoolType == 4 )//VHSC2
                {
                    str.append("Total 9 A+");
                }
                else if( m_scoolType == 1 || m_scoolType == 2 ) //HSS
                {
                    str.append("Total 5 A+");
                }
                else if( m_scoolType == 0 )//HighSchool
                {
                    str.append("Total 9 A+");
                }


                cellL->dynamicCall("SetValue(const QVariant&)",QVariant(str));

                M ="M"+QString::number(4);
                cellM = sheet->querySubObject("Range(QVariant, QVariant)",M);
                cellM->dynamicCall("SetValue(const QVariant&)",QVariant(m_APlus9Count));

                L ="L"+QString::number(5);
                cellL = sheet->querySubObject("Range(QVariant, QVariant)",L);
                cellL->dynamicCall("SetValue(const QVariant&)",QVariant("Total Pass"));

                M ="M"+QString::number(5);
                cellM = sheet->querySubObject("Range(QVariant, QVariant)",M);
                cellM->dynamicCall("SetValue(const QVariant&)",QVariant(m_Pass));

                L ="L"+QString::number(6);
                cellL = sheet->querySubObject("Range(QVariant, QVariant)",L);
                cellL->dynamicCall("SetValue(const QVariant&)",QVariant("Total Fail"));

                M ="M"+QString::number(6);
                cellM = sheet->querySubObject("Range(QVariant, QVariant)",M);
                cellM->dynamicCall("SetValue(const QVariant&)",QVariant(m_Fail));

                L ="L"+QString::number(7);
                cellL = sheet->querySubObject("Range(QVariant, QVariant)",L);
                cellL->dynamicCall("SetValue(const QVariant&)",QVariant("Percentage of Pass"));

                M ="M"+QString::number(7);
                cellM = sheet->querySubObject("Range(QVariant, QVariant)",M);
                float diff = m_Pass-m_Fail;
                cellM->dynamicCall("SetValue(const QVariant&)",QVariant((diff/m_Pass)*100));
            }
        }}




    workbook->dynamicCall("SaveAs(const QString&)",m_filePath);
    QThread::msleep(300);
    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;

    emit progress(false);
}

void SchoolModel::vhsc()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* workbook = workbooks->querySubObject( "Open(const QString&)", m_filePath );
    QAxObject* sheets = workbook->querySubObject( "Worksheets" );

    //worksheets count
    int count = sheets->dynamicCall("Count()").toInt();

    count = sheets->property("Count").toInt();
    qDebug() << "Count" << count;
    m_rankInfo.clear();


    for (int i=1; i<=1; i++) //cycle through sheets
    {
        //sheet pointer
        QAxObject* sheet = sheets->querySubObject( "Item( int )", i );

        QAxObject* rows = sheet->querySubObject( "Rows" );
        int rowCount = rows->dynamicCall( "Count()" ).toInt(); //unfortunately, always returns 255, so you have to check somehow validity of cell values
        // qDebug() << "Count" << rowCount;
        QAxObject* columns = sheet->querySubObject( "Columns" );
        // int columnCount = columns->property("Count").toInt();
        QAxObject *cellB, *cellC;
        //VHSCInfo  vhscInfo;

        int pointsOnStudent = 0;
        RankInfo studentDetails;

        int subjectCount =0;

        for (int row=1; row <= rowCount; row++)
        {
            QString B ="A"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString rollNumber = cellB->dynamicCall("Value").toString();
            qDebug() << rollNumber;
            if(!rollNumber.isEmpty())
                studentDetails.rollNumber = rollNumber;


            B ="B"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString name = cellB->dynamicCall("Value").toString();
            qDebug() << name;
            if( !name.isEmpty())
            {
                studentDetails.name = name;
            }


            B ="D"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;

            pointsOnStudent += m_gradeMap[grade];

            QString C ="C"+QString::number(row);
            cellC = sheet->querySubObject("Range(QVariant, QVariant)",C);
            QString subject = cellC->dynamicCall("Value").toString();
            qDebug() << subject;
            if(!subject.contains("Maths"))
            {



                if( grade == "A+")
                {
                    studentDetails.gradeMap[ "A+"] += 1 ;
                }
                else  if( grade == "A")
                {
                    studentDetails.gradeMap[ "A"] += 1 ;
                }
                else  if( grade == "B+")
                {
                    studentDetails.gradeMap[ "B+"] += 1 ;
                }
                else  if( grade == "B")
                {
                    studentDetails.gradeMap[ "B"] += 1 ;
                }
                else  if( grade == "C+")
                {
                    studentDetails.gradeMap[ "C+"] += 1 ;
                }
                else  if( grade == "C")
                {
                    studentDetails.gradeMap[ "C"] += 1 ;
                }
                else  if( grade == "D+")
                {
                    studentDetails.gradeMap[ "D+"] += 1 ;
                }
                else  if( grade == "D")
                {
                    studentDetails.gradeMap[ "D"] += 1 ;
                }
                else  if( grade == "E+")
                {
                    studentDetails.gradeMap[ "E+"] += 1 ;
                }
                else  if( grade == "E")
                {
                    studentDetails.gradeMap[ "E"] += 1 ;
                }
            }

            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["A+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["A"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["B+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["B"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["C+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["C"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["D+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["D"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["E+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["E"]));


            B ="E"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString  passFail = cellB->dynamicCall("Value").toString();
            qDebug() <<  passFail;

            if( !passFail.isEmpty())
                studentDetails.passFail =  passFail;

            if(studentDetails.passFail != "EHS" && !studentDetails.passFail.isEmpty())
            {
                subjectCount+=11;
                row+=10;
                studentDetails = {};
                pointsOnStudent = 0;
                m_Fail +=1;
                continue;
            }

            subjectCount++;
            if( subjectCount%11 == 0  )
            {

                if( pointsOnStudent == 0 )
                    break;

                m_Pass+=1;
                if( m_rankInfo.contains(pointsOnStudent) )
                {
                    QMultiMap<int, QMultiMap<unsigned long long, RankInfo>>::iterator itr = m_rankInfo.find(pointsOnStudent);
                    unsigned long long rank = studentDetails.m_rank.toLongLong();

                    QMultiMap<unsigned long long, RankInfo>& map = itr.value();;
                    // map.insert()
                    map.insert(rank,studentDetails);


                }
                else
                {
                    QMultiMap<unsigned long long, RankInfo> temp;
                    unsigned long long rank = studentDetails.m_rank.toLongLong();
                    temp.insert(rank,studentDetails);
                    m_rankInfo.insert(pointsOnStudent,temp);
                }

                pointsOnStudent = 0;
                if( studentDetails.gradeMap["A+"] == 10 )
                {
                    m_APlus10Count++;
                }
                else if( studentDetails.gradeMap["A+"] == 9 )
                {
                    m_APlus9Count++;
                }
                studentDetails = {};
            }
        }
    }

    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;
    reportFullAPlus();

}

void SchoolModel::vhsc1()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* workbook = workbooks->querySubObject( "Open(const QString&)", m_filePath );
    QAxObject* sheets = workbook->querySubObject( "Worksheets" );

    //worksheets count
    int count = sheets->dynamicCall("Count()").toInt();

    count = sheets->property("Count").toInt();
    qDebug() << "Count" << count;
    m_rankInfo.clear();


    for (int i=1; i<=1; i++) //cycle through sheets
    {
        //sheet pointer
        QAxObject* sheet = sheets->querySubObject( "Item( int )", i );

        QAxObject* rows = sheet->querySubObject( "Rows" );
        int rowCount = rows->dynamicCall( "Count()" ).toInt(); //unfortunately, always returns 255, so you have to check somehow validity of cell values
        // qDebug() << "Count" << rowCount;
        QAxObject* columns = sheet->querySubObject( "Columns" );
        // int columnCount = columns->property("Count").toInt();
        QAxObject *cellB, *cellC;
        //VHSCInfo  vhscInfo;

        int pointsOnStudent = 0;
        RankInfo studentDetails;
        studentDetails = {};
        int subjectCount =0;

        for (int row=1; row <= rowCount; row++)
        {
            QString B ="A"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString rollNumber = cellB->dynamicCall("Value").toString();
            qDebug() << rollNumber;
            if(!rollNumber.isEmpty())
                studentDetails.rollNumber = rollNumber;


            B ="B"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString name = cellB->dynamicCall("Value").toString();
            qDebug() << name;
            if( !name.isEmpty())
            {
                studentDetails.name = name;
            }


            B ="E"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;

            pointsOnStudent += m_gradeMap[grade];

            studentDetails.mark += grade.toInt();

            QString C ="C"+QString::number(row);
            cellC = sheet->querySubObject("Range(QVariant, QVariant)",C);
            QString subject = cellC->dynamicCall("Value").toString();
            qDebug() << subject;


            int modulo = row%8;

                if( modulo == 1 )
                {
                    studentDetails.gradeMap[ "A+"] += grade.toInt() ;
                }
                else  if( modulo == 2)
                {
                    studentDetails.gradeMap[ "A"] += grade.toInt() ;
                }
                else  if( modulo == 3)
                {
                    studentDetails.gradeMap[ "B+"] += grade.toInt() ;
                }
                else  if( modulo == 4)
                {
                    studentDetails.gradeMap[ "B"] += grade.toInt() ;
                }
                else  if( modulo == 5)
                {
                    studentDetails.gradeMap[ "C+"] += grade.toInt() ;
                }
                else  if( modulo == 6)
                {
                    studentDetails.gradeMap[ "C"] += grade.toInt() ;
                }
                else  if( modulo == 7)
                {
                    studentDetails.gradeMap[ "D+"] += grade.toInt() ;
                }
                else  if( modulo == 0 )
                {
                    studentDetails.gradeMap[ "D"] += grade.toInt() ;
                }




            subjectCount++;
            if( subjectCount%8 == 0  )
            {


                if( subject.isEmpty() )
                    break;
                studentDetails.m_rank = QString::number(studentDetails.mark);
                m_Pass+=1;
                if( m_rankInfo.contains(pointsOnStudent) )
                {
                    QMultiMap<int, QMultiMap<unsigned long long, RankInfo>>::iterator itr = m_rankInfo.find(pointsOnStudent);
                    unsigned long long rank = studentDetails.m_rank.toLongLong();

                    QMultiMap<unsigned long long, RankInfo>& map = itr.value();;
                    // map.insert()
                    map.insert(rank,studentDetails);
                }
                else
                {
                    QMultiMap<unsigned long long, RankInfo> temp;
                    unsigned long long rank = studentDetails.m_rank.toLongLong();
                    temp.insert(rank,studentDetails);
                    m_rankInfo.insert(pointsOnStudent,temp);
                }

                pointsOnStudent = 0;
                studentDetails = {};
            }
        }
    }

    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;
    reportFullAPlus();
}

void SchoolModel::hss()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* workbook = workbooks->querySubObject( "Open(const QString&)", m_filePath );
    QAxObject* sheets = workbook->querySubObject( "Worksheets" );



    //worksheets count
    int count = sheets->dynamicCall("Count()").toInt();

    count = sheets->property("Count").toInt();
    qDebug() << "Count" << count;
    m_rankInfo.clear();
    for (int i=1; i<=1; i++) //cycle through sheets
    {
        //sheet pointer
        QAxObject* sheet = sheets->querySubObject( "Item( int )", i );

        QAxObject* rows = sheet->querySubObject( "Rows" );
        int rowCount = rows->dynamicCall( "Count()" ).toInt(); //unfortunately, always returns 255, so you have to check somehow validity of cell values
        QAxObject *cellB, *cellC;
        //VHSCInfo  vhscInfo;

        int pointsOnStudent = 0;
        RankInfo studentDetails;

        int previousRollNumber = 0;

        bool isFirst = true;
        for (int row=1; row <= rowCount; row++)
        {
            QString B ="A"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString rollNumber = cellB->dynamicCall("Value").toString();
            qDebug() << rollNumber;
            if( (qAbs(previousRollNumber-rollNumber.toInt()) > 10) && !isFirst )
            {
                qDebug() << "skipped" << rollNumber;
                isFirst = false;
                continue;
            }
            if(!rollNumber.isEmpty())
            {
                previousRollNumber = rollNumber.toInt();
                studentDetails.rollNumber = rollNumber;
            }
            else
            {
                break;
            }

            B ="B"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString name = cellB->dynamicCall("Value").toString();
            qDebug() << name;
            if( !name.isEmpty())
            {
                studentDetails.name = name;
            }

            QStringList list;
            B ="C"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString coarce = cellB->dynamicCall("Value").toString();
            qDebug() << coarce;



            B ="E"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="G"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="I"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="K"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="M"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="O"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);


            foreach (QString strGrade, list) {

                QStringList list = strGrade.split(" ");
                QString grade =  list.value(0);
                QString strMark =  list.value(1);

                QString markCorrected = strMark.replace(")", "");
                QString mark = markCorrected.replace("(", "");
                studentDetails.mark += mark.toInt();

                qDebug() << "....................................."<<  studentDetails.mark << grade;

                pointsOnStudent += m_gradeMap[grade];
                if( grade == "A+")
                {
                    studentDetails.gradeMap[ "A+"] += 1 ;
                }
                else  if( grade == "A")
                {
                    studentDetails.gradeMap[ "A"] += 1 ;
                }
                else  if( grade == "B+")
                {
                    studentDetails.gradeMap[ "B+"] += 1 ;
                }
                else  if( grade == "B")
                {
                    studentDetails.gradeMap[ "B"] += 1 ;
                }
                else  if( grade == "C+")
                {
                    studentDetails.gradeMap[ "C+"] += 1 ;
                }
                else  if( grade == "C")
                {
                    studentDetails.gradeMap[ "C"] += 1 ;
                }
                else  if( grade == "D+")
                {
                    studentDetails.gradeMap[ "D+"] += 1 ;
                }
                else  if( grade == "D")
                {
                    studentDetails.gradeMap[ "D"] += 1 ;
                }
                else  if( grade == "E+")
                {
                    studentDetails.gradeMap[ "E+"] += 1 ;
                }
                else  if( grade == "E")
                {
                    studentDetails.gradeMap[ "E"] += 1 ;
                }
            }

            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["A+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["A"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["B+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["B"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["C+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["C"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["D+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["D"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["E+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["E"]));


            B ="P"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString  passFail = cellB->dynamicCall("Value").toString();
            qDebug() <<  passFail;

            if( !passFail.isEmpty())
                studentDetails.passFail =  passFail;

            if( studentDetails.passFail == "EHS" )
            {
                m_Pass+=1;
            }
            else
            {
                m_Fail +=1;
            }

            if(  !studentDetails.passFail.isEmpty() && studentDetails.passFail != "NHS" )
            {
                if( m_rankInfo.contains(pointsOnStudent) )
                {
                    QMultiMap<int, QMultiMap<unsigned long long, RankInfo>>::iterator itr = m_rankInfo.find(pointsOnStudent);
                    unsigned long long rank = studentDetails.m_rank.toLongLong();

                    QMultiMap<unsigned long long, RankInfo>& map = itr.value();;
                    // map.insert()
                    map.insert(rank,studentDetails);


                }
                else
                {
                    QMultiMap<unsigned long long, RankInfo> temp;
                    unsigned long long rank = studentDetails.m_rank.toLongLong();
                    temp.insert(rank,studentDetails);
                    m_rankInfo.insert(pointsOnStudent,temp);
                }

                if( studentDetails.gradeMap["A+"] == 6 )
                {
                    m_APlus10Count++;
                }
                else if( studentDetails.gradeMap["A+"] == 5 )
                {
                    m_APlus9Count++;
                }
            }
            pointsOnStudent = 0;
            studentDetails = {};
        }
    }

    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;
    reportFullAPlus();

}

void SchoolModel::hss1()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* workbook = workbooks->querySubObject( "Open(const QString&)", m_filePath );
    QAxObject* sheets = workbook->querySubObject( "Worksheets" );



    //worksheets count
    int count = sheets->dynamicCall("Count()").toInt();

    count = sheets->property("Count").toInt();
    qDebug() << "Count" << count;
    m_rankInfo.clear();
    for (int i=1; i<=1; i++) //cycle through sheets
    {
        //sheet pointer
        QAxObject* sheet = sheets->querySubObject( "Item( int )", i );

        QAxObject* rows = sheet->querySubObject( "Rows" );
        int rowCount = rows->dynamicCall( "Count()" ).toInt(); //unfortunately, always returns 255, so you have to check somehow validity of cell values
        QAxObject *cellB, *cellC;
        //VHSCInfo  vhscInfo;

        int pointsOnStudent = 0;
        RankInfo studentDetails;
        studentDetails.mark = 0;
        studentDetails.m_rank = "";

        int previousRollNumber = 0;

        bool isFirst = true;
        for (int row=1; row <= rowCount; row++)
        {
            QString B ="A"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString rollNumber = cellB->dynamicCall("Value").toString();
            qDebug() << rollNumber;
            if( (qAbs(previousRollNumber-rollNumber.toInt()) > 10) && !isFirst )
            {
                qDebug() << "skipped" << rollNumber;
                isFirst = false;
                continue;
            }
            if(!rollNumber.isEmpty())
            {
                previousRollNumber = rollNumber.toInt();
                studentDetails.rollNumber = rollNumber;
            }
            else
            {
                break;
            }

            B ="B"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString name = cellB->dynamicCall("Value").toString();
            qDebug() << name;
            if( !name.isEmpty())
            {
                studentDetails.name = name;
            }

            QStringList list;
            B ="C"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString coarce = cellB->dynamicCall("Value").toString();
            qDebug() << coarce;



            B ="E"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="G"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="I"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="K"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="M"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="O"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);


            int i = 0;
            foreach (QString strGrade, list) {

                //QString markCorrected = strMark.replace(")", "");
                QString mark = strGrade;
                studentDetails.mark += mark.toInt();

                qDebug() << "....................................."<<  studentDetails.mark << grade;

                pointsOnStudent += m_gradeMap[grade];
                if( i == 0)
                {
                    studentDetails.gradeMap[ "A+"] +=  mark.toInt();
                }
                else  if( i == 1 )
                {
                    studentDetails.gradeMap[ "A"] +=  mark.toInt();
                }
                else  if( i == 2 )
                {
                    studentDetails.gradeMap[ "B+"] += mark.toInt();
                }
                else  if( i == 3 )
                {
                    studentDetails.gradeMap[ "B"] += mark.toInt();
                }
                else  if( i == 4 )
                {
                    studentDetails.gradeMap[ "C+"] += mark.toInt();
                }
                else  if( i == 5 )
                {
                    studentDetails.gradeMap[ "C"] += mark.toInt();
                }
                else  if( i == 6 )
                {
                    studentDetails.gradeMap[ "D+"] += mark.toInt();
                }
                else  if( i == 7 )
                {
                    studentDetails.gradeMap[ "D"] += mark.toInt();
                }
                else  if( i == 8 )
                {
                    studentDetails.gradeMap[ "E+"] += mark.toInt();
                }
                else  if( i == 9 )
                {
                    studentDetails.gradeMap[ "E"] += mark.toInt();
                }
                i++;
            }

            studentDetails.m_rank = QString::number(studentDetails.mark);

            B ="P"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString  passFail = cellB->dynamicCall("Value").toString();
            qDebug() <<  passFail;

            if( !passFail.isEmpty())
                studentDetails.passFail =  passFail;

            if( studentDetails.mark != 0 )
            {
                m_Pass+=1;
            }
            else
            {
                m_Fail +=1;
            }

            if( studentDetails.mark != 0 )
            {
                if( m_rankInfo.contains(pointsOnStudent) )
                {
                    QMultiMap<int, QMultiMap<unsigned long long, RankInfo>>::iterator itr = m_rankInfo.find(pointsOnStudent);
                    unsigned long long rank = studentDetails.m_rank.toLongLong();

                    QMultiMap<unsigned long long, RankInfo>& map = itr.value();;
                    // map.insert()
                    map.insert(rank,studentDetails);


                }
                else
                {
                    QMultiMap<unsigned long long, RankInfo> temp;
                    unsigned long long rank = studentDetails.m_rank.toLongLong();
                    temp.insert(rank,studentDetails);
                    m_rankInfo.insert(pointsOnStudent,temp);
                }

                if( studentDetails.gradeMap["A+"] == 6 )
                {
                    m_APlus10Count++;
                }
                else if( studentDetails.gradeMap["A+"] == 5 )
                {
                    m_APlus9Count++;
                }
            }
            pointsOnStudent = 0;
            studentDetails = {};
        }
    }

    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;
    reportFullAPlus();

}


void SchoolModel::highSchool()
{
    QAxObject* excel = new QAxObject( "Excel.Application", 0 );
    QAxObject* workbooks = excel->querySubObject( "Workbooks" );
    QAxObject* workbook = workbooks->querySubObject( "Open(const QString&)", m_filePath );
    QAxObject* sheets = workbook->querySubObject( "Worksheets" );

    //worksheets count
    int count = sheets->dynamicCall("Count()").toInt();

    count = sheets->property("Count").toInt();
    qDebug() << "Count" << count;
    m_rankInfo.clear();
    for (int i=1; i<=1; i++) //cycle through sheets
    {
        //sheet pointer
        QAxObject* sheet = sheets->querySubObject( "Item( int )", i );

        QAxObject* rows = sheet->querySubObject( "Rows" );
        int rowCount = rows->dynamicCall( "Count()" ).toInt(); //unfortunately, always returns 255, so you have to check somehow validity of cell values
        // qDebug() << "Count" << rowCount;
        QAxObject* columns = sheet->querySubObject( "Columns" );
        // int columnCount = columns->property("Count").toInt();
        QAxObject *cellB, *cellC;
        //VHSCInfo  vhscInfo;

        int pointsOnStudent = 0;
        RankInfo studentDetails;

        for (int row=1; row <= rowCount; row++)
        {



            QString B ="A"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString rollNumber = cellB->dynamicCall("Value").toString();
            qDebug() << rollNumber;
            if(!rollNumber.isEmpty())
            {
                studentDetails.rollNumber = rollNumber;
            }
            else
            {
                break;
            }


            B ="B"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString name = cellB->dynamicCall("Value").toString();
            qDebug() << name;
            if( !name.isEmpty())
            {
                studentDetails.name = name;
            }

            QStringList list;
            B ="C"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString grade = cellB->dynamicCall("Value").toString();
            //qDebug() << grade;
            list.append(grade);


            B ="D"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            //qDebug() << grade;
            list.append(grade);

            B ="E"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="F"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="G"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="H"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="I"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="J"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="K"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            B ="L"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            grade = cellB->dynamicCall("Value").toString();
            qDebug() << grade;
            list.append(grade);

            foreach (QString strGrade, list) {



                pointsOnStudent += m_gradeMap[strGrade];
                if( strGrade == "A+")
                {
                    studentDetails.gradeMap[ "A+"] += 1 ;
                }
                else  if( strGrade == "A")
                {
                    studentDetails.gradeMap[ "A"] += 1 ;
                }
                else  if( strGrade == "B+")
                {
                    studentDetails.gradeMap[ "B+"] += 1 ;
                }
                else  if( strGrade == "B")
                {
                    studentDetails.gradeMap[ "B"] += 1 ;
                }
                else  if( strGrade == "C+")
                {
                    studentDetails.gradeMap[ "C+"] += 1 ;
                }
                else  if( strGrade == "C")
                {
                    studentDetails.gradeMap[ "C"] += 1 ;
                }
                else  if( strGrade == "D+")
                {
                    studentDetails.gradeMap[ "D+"] += 1 ;
                }
                else  if( strGrade == "D")
                {
                    studentDetails.gradeMap[ "D"] += 1 ;
                }
                else  if( strGrade == "E+")
                {
                    studentDetails.gradeMap[ "E+"] += 1 ;
                }
                else  if( strGrade == "E")
                {
                    studentDetails.gradeMap[ "E"] += 1 ;
                }
            }


            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["A+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["A"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["B+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["B"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["C+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["C"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["D+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["D"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["E+"]));
            studentDetails.m_rank.append(QString::number(studentDetails.gradeMap["E"]));


            B ="M"+QString::number(row);
            cellB = sheet->querySubObject("Range(QVariant, QVariant)",B);
            QString  passFail = cellB->dynamicCall("Value").toString();
            qDebug() <<  passFail;

            if( !passFail.isEmpty())
                studentDetails.passFail =  passFail;

            if(  studentDetails.passFail == "EHS" )
            {
                m_Pass+=1;
                if( m_rankInfo.contains(pointsOnStudent) )
                {
                    QMultiMap<int, QMultiMap<unsigned long long, RankInfo>>::iterator itr = m_rankInfo.find(pointsOnStudent);
                    unsigned long long rank = studentDetails.m_rank.toLongLong();

                    QMultiMap<unsigned long long, RankInfo>& map = itr.value();;
                    // map.insert()
                    map.insert(rank,studentDetails);


                }
                else
                {
                    QMultiMap<unsigned long long, RankInfo> temp;
                    unsigned long long rank = studentDetails.m_rank.toLongLong();
                    temp.insert(rank,studentDetails);
                    m_rankInfo.insert(pointsOnStudent,temp);
                }

                if( studentDetails.gradeMap["A+"] == 10 )
                {
                    m_APlus10Count++;
                }
                else if( studentDetails.gradeMap["A+"] == 9 )
                {
                    m_APlus9Count++;
                }
            }
            else
            {
                m_Fail +=1;
            }
            pointsOnStudent = 0;
            studentDetails = {};
        }
    }

    workbook->dynamicCall("Close()");
    excel->dynamicCall("Quit()");
    delete excel;
    reportFullAPlus();

}



