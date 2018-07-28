#ifndef SCHOOLMODE_H
#define SCHOOLMODE_H

#include <QObject>
#include <QMap>

struct Subject
{
    QString mark;
    QString grade;
};
struct VHSCInfo
{
    QString name;
    QMap<QString, Subject> subjectMap;
    QString passFaile;
};

struct RankInfo
{
      QString name;
      QMap<QString, int> gradeMap;
      QString rollNumber;
      QString passFail;
      int mark;
      QString m_rank;
};

class SchoolModel : public QObject
{
    Q_OBJECT
public:
     SchoolModel();
     Q_INVOKABLE void setFilePath( QString );
     Q_INVOKABLE void generateReport(int type);

private:
     void reportFullAPlus();
     void vhsc();
     void vhsc1();
     void hss1();
     void hss();
     void highSchool();

signals:

     void progress( bool isProgress );
     void sheetIsNotAvalable();

public slots:

private:
     QString m_filePath;
     int m_scoolType;
     QMap<int,VHSCInfo> vhscInfoMap;
     QMap<QString, int> m_gradeMap;
     QMultiMap<int,QMultiMap<unsigned long long, RankInfo>> m_rankInfo;



     int m_APlus10Count=0;
     int m_APlus9Count=0;
     int m_Pass = 0;
     int m_Fail = 0;

};

#endif // SCHOOLMODE_H
