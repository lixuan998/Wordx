#pragma once
#include <QString>
#include <string>
#include <vector>
#include <map>
#include <opencv2/core/core.hpp>
#include <iostream>

class Info
{
    public:
        Info();
        Info(std::map<QString, QString> label_to_text_map, std::map<QString, cv::Mat> label_to_mat_map);
        ~Info();
        void addInfo(QString label, QString text);
        void addInfo(QString label, cv::Mat image);
        void clear();
        std::map<QString, QString> getLabelText();
        std::map<QString, cv::Mat> getLabelImage();
    private:
        std::map<QString, QString> label_to_text_map;
        std::map<QString, cv::Mat> label_to_mat_map;
};