#include "info.h"

/*---------------
Public Functions
---------------*/

Info::Info()
{
    //std::cout << "an instance of Info has been created" << std::endl;
}

Info::Info(std::map<QString, QString> label_to_text_map, std::map<QString, cv::Mat> label_to_mat_map)
{
    this->label_to_text_map = label_to_text_map;
    this->label_to_mat_map = label_to_mat_map;
}

Info::~Info()
{
    //std::cout << "an instance of Info has been distroyed" << std::endl;
}

void Info::clear()
{
    label_to_text_map.clear();
    label_to_mat_map.clear();
}

void Info::addInfo(QString label, QString text)
{
    label.replace("<", "&lt;");
    label.replace(">", "&gt;");
    label_to_text_map[label] = text;
}

void Info::addInfo(QString label, cv::Mat image)
{
    label.replace("<", "&lt;");
    label.replace(">", "&gt;");
    label_to_mat_map[label] = image;
}

std::map<QString, QString> Info::getLabelText()
{
    return label_to_text_map;
}

std::map<QString, cv::Mat> Info::getLabelImage()
{
    return label_to_mat_map;
}
