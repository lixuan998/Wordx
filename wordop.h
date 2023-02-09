#ifndef __WORDOP_H__
#define __WORDOP_H__

#define SINGLE single
#define DASHED dashed
#define DASHSMALLGAP dashSmallGap

#define UP 0x0001
#define DOWN 0x0002
#include "fileop.h"
#include <fstream>
#include <vector>
#include <QDir>
#include <QFile>

class WordOp{
    public:
        WordOp(std::string filepath = "");

        void open(std::string filepath = "");

        void close();

        int replaceText(std::string mark, std::string replaced_text, std::string &rep_str);

        int replaceText(std::vector<std::string> marks, std::vector<std::string> replaced_texts, std::string &rep_str);

        int replaceText(std::string mark_file, std::string replaced_text_file);

        int replaceImage(std::string mark, std::string replace_image);

        int replaceImageFromFile(std::string mark_file_path, std::string replace_image_file_path);

        int addInfoRecursive(std::vector<int> indexs, std::vector<std::string> info_file_path);

        int addInfoRecursive(std::vector<int> indexs, std::vector<std::vector<std::vector<std::string>>> rep_infos);

        int addTableRows(std::vector<int> indexs, std::vector<std::string> info_file_path);

    private:
        int addImage(std::string mark, std::string replaced_image);
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
};

#endif
