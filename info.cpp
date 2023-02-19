#include "info.h"

Info::Info()
{
    //std::cout << "an instance of Info has been created" << std::endl;
}

Info::Info(std::map<std::string, std::string> label_to_text_map, std::map<std::string, cv::Mat> label_to_mat_map)
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

void Info::addInfo(std::string label, std::string text)
{
    label_to_text_map[label] = text;
}

void Info::addInfo(std::string label, cv::Mat image)
{
    label_to_mat_map[label] = image;
}

std::map<std::string, std::string> Info::getLabelText()
{
    return label_to_text_map;
}

std::map<std::string, cv::Mat> Info::getLabelImage()
{
    return label_to_mat_map;
}
