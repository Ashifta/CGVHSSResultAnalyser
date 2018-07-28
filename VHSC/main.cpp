#include <QGuiApplication>
#include <QQmlApplicationEngine>
#include <QQmlContext>

#include "schoolmodel.h"
int main(int argc, char *argv[])
{
    QGuiApplication app(argc, argv);

    QQmlApplicationEngine engine;

    SchoolModel model;
    engine.rootContext()->setContextProperty("schoolModel", &model);
    engine.load(QUrl(QStringLiteral("qrc:/main.qml")));

    return app.exec();
}
