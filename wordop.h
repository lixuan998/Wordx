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



#define LP_START_MARK "${LP_START}"
#define LP_END_MARK "${LP_END}"

#define TABLE_LOOP_START_MARK "${TB_START}"
#define TABLE_LOOP_END_MARK "${TB_END}"

#include <vector>
#include <map>
#include <QDir>
#include <QImage>
#include <QFile>
#include <QDebug>
#include <QCoreApplication>
#include <QThreadPool>

#include <opencv2/imgcodecs.hpp>
#include <opencv2/core/core.hpp>

#include "fileop.h"
#include "info.h"
#include "add_image_thread.h"

#define IMAGE_MODEL_PATH (QCoreApplication::applicationDirPath() + "/xml_models/image.xml")

class WordOp
{
    public:
        WordOp(QString filepath = "");

        WordOp(QString filepath, QString des_path);
        ~WordOp();

        void open(QString filepath = "");

        void open(QString filepath, QString des_path);

        void close();

        int replaceText(QString mark, QString replaced_text, QString &rep_str);
        int replaceText(Info info);
        int replaceText(std::vector<QString> marks, std::vector<QString> replaced_texts, QString &rep_str);
        int replaceText(std::vector<QString> marks, std::vector<QString> replaced_texts);
        int replaceText(QString mark_file, QString replaced_text_file);

        int replaceImage(std::vector<QString> marks, std::vector<QString> replace_image_paths);
        int replaceImageFromMat(std::vector<QString> marks, std::vector<cv::Mat> replace_mat_images);
        int replaceImageFromFile(QString mark_file_path, QString replace_image_file_path);

        int addInfoRecursive(std::vector<int> indexs, std::vector<Info> &infos);
        
        int addTableRows(std::vector<int> indexs, std::vector<Info> &infos);
    
    private:
        int addImage(QString replace_image_path);
        int addImageFromMat(cv::Mat replace_mat_image);
        void readXml(QString &xml_file, QString filepath = "");
        void readList(std::vector<std::vector<QString>> &str_list, QString filepath = "");
        void writeXml(QString &xml_file, QString filepath = "");
        void analyzeXml(std::vector<QString> &analysis_xml, QString origin_xml, QString flag);
        int myFind(QString src, QString des, int index = 1);
        int findAround(QString src, QString des, int cur_pos, int direction);

    public:
        static QString line_feed_head, line_feed_tail;

    private:
        QString filepath;
        QString des_path;
        QString cache_path;
        QString document_xml, style_xml, doc_rels_xml;
        QString image_rels;
        QString origin_work_path;
        QImage image;
        Add_Image_Thread *add_image_thread;
        int image_sn;
};
#endif
