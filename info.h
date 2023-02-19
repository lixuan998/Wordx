#pragma once
#include <string>
#include <vector>
#include <map>
#include <opencv2/core/core.hpp>
#include <iostream>

class Info
{
    public:
        Info();
        Info(std::map<std::string, std::string> label_to_text_map, std::map<std::string, cv::Mat> label_to_mat_map);
        ~Info();
        void addInfo(std::string label, std::string text);
        void addInfo(std::string label, cv::Mat image);
        void clear();
        std::map<std::string, std::string> getLabelText();
        std::map<std::string, cv::Mat> getLabelImage();
    private:
        std::map<std::string, std::string> label_to_text_map;
        std::map<std::string, cv::Mat> label_to_mat_map;
};