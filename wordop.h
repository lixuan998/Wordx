/*********************************************************************************************************************************************
File name:  wordop.h
Author: Li Xuan
Version: 1.0
Date:  2023-2-19
Description:  This module is used to modify files with suffix ".docx"
Function List: 
    

***********************************************************************************************************************************************/


#ifndef __WORDOP_H__
#define __WORDOP_H__

#define SINGLE single
#define DASHED dashed
#define DASHSMALLGAP dashSmallGap

#define UP 0x0001
#define DOWN 0x0002

#include <fstream>
#include <vector>
#include <map>
#include <QDir>
#include <QImage>
#include <QFile>

#include <opencv2/imgcodecs.hpp>
#include <opencv2/core/core.hpp>

#include "fileop.h"
#include "info.h"

class WordOp
{
    public:
        WordOp(std::string filepath = "");
        ~WordOp();

        void open(std::string filepath = "");
        void close();

        int replaceText(std::string mark, std::string replaced_text, std::string &rep_str);
        int replaceText(std::vector<std::string> marks, std::vector<std::string> replaced_texts, std::string &rep_str);
        int replaceText(std::string mark_file, std::string replaced_text_file);

        int replaceImage(std::vector<std::string> marks, std::vector<std::string> replace_image_paths);
        int replaceImageFromMat(std::vector<std::string> marks, std::vector<cv::Mat> replace_mat_images);
        int replaceImageFromFile(std::string mark_file_path, std::string replace_image_file_path);

        int addInfoRecursive(std::vector<int> indexs, std::vector<Info> &infos);
        
        int addTableRows(std::vector<int> indexs, std::vector<std::string> info_file_path);
        int addTableRows(std::vector<int> indexs, std::vector<Info> &infos);

    private:
        int addImage(std::string replace_image_path);
        int addImageFromMat(cv::Mat replace_mat_image);
        void readXml(std::string &xml_file, std::string filepath = "");
        void readList(std::vector<std::vector<std::string>> &str_list, std::string filepath = "");
        void writeXml(std::string &xml_file, std::string filepath = "");
        void analyzeXml(std::vector<std::string> &analysis_xml, std::string origin_xml, std::string flag);
        int myFind(std::string src, std::string des, int index = 1);
        int findAround(std::string src, std::string des, int cur_pos, int direction);

    private:
        std::string filepath;
        std::string cache_path;
        std::string document_xml, style_xml, doc_rels_xml;
        std::string image_rels;
        QImage image;
};
#endif
