#ifndef __WORDOP_H__
#define __WORDOP_H__

#define SINGLE single
#define DASHED dashed
#define DASHSMALLGAP dashSmallGap
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

        int replaceImage(std::string mark, std::string replaced_image);

        int addInfoRecursive(int index, std::string info_file_path);

        int createTable(int pos, int rows, int cols, int row_height, std::string border_line_type, std::string border_line_color, std::string text_color, std::string text);

        int createTable(int pos, int rows, int cols, std::vector< std::vector<std::string> > border_line_type, std::vector< std::vector<std::string> > border_line_colors, std::vector<std::string> text_color, std::vector<std::string> text);

    private:
        void readXml(std::string &xml_file, std::string filepath = "");
        void readList(std::vector<std::vector<std::string>> &str_list, std::string filepath = "");
        void writeXml(std::string &xml_file, std::string filepath = "");
        void analyzeXml(std::vector<std::string> &analysis_xml, std::string origin_xml);
        int myFind(std::string src, std::string des, int index = 1);

        std::vector<int> findImgModelPos(std::string src, int index);
    private:
        std::string filepath;
        std::string cache_path;
        std::string document_xml, style_xml, doc_rels_xml;
        std::string image_rels;
};

#endif
