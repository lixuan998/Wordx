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

#define NORMAL_IMAGE_SIZE_TIMES 3000
#define TABLE_IMAGE_SIZE_TIMES 700

#include <fstream>
#include <vector>
#include <map>
#include <QDir>
#include <QImage>
#include <QFile>
#include <QDebug>

#include <opencv2/imgcodecs.hpp>
#include <opencv2/core/core.hpp>

#include "fileop.h"
#include "info.h"

class WordOp
{
    public:
        WordOp(QString filepath = "");
        ~WordOp();

        void open(QString filepath = "");
        void close();

        int replaceText(QString mark, QString replaced_text, QString &rep_str);
        int replaceText(std::vector<QString> marks, std::vector<QString> replaced_texts, QString &rep_str);
        int replaceText(std::vector<QString> marks, std::vector<QString> replaced_texts);
        int replaceText(QString mark_file, QString replaced_text_file);

        int replaceImage(std::vector<QString> marks, std::vector<QString> replace_image_paths);
        int replaceImageFromMat(std::vector<QString> marks, std::vector<cv::Mat> replace_mat_images);
        int replaceImageFromFile(QString mark_file_path, QString replace_image_file_path);

        int addInfoRecursive(std::vector<int> indexs, std::vector<Info> &infos);
        
        int addTableRows(std::vector<int> indexs, std::vector<QString> info_file_path);
        int addTableRows(std::vector<int> indexs, std::vector<Info> &infos);
    
    public:
        static QString line_feed_head, line_feed_tail;

    private:
        int addImage(QString replace_image_path);
        int addImageFromMat(cv::Mat replace_mat_image);
        void readXml(QString &xml_file, QString filepath = "");
        void readList(std::vector<std::vector<QString>> &str_list, QString filepath = "");
        void writeXml(QString &xml_file, QString filepath = "");
        void analyzeXml(std::vector<QString> &analysis_xml, QString origin_xml, QString flag);
        int myFind(QString src, QString des, int index = 1);
        int findAround(QString src, QString des, int cur_pos, int direction);

    private:
        QString filepath;
        QString cache_path;
        QString document_xml, style_xml, doc_rels_xml;
        QString image_rels;
        
        QImage image;
};
#endif
