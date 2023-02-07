#ifndef __WORDOP_H__
#define __WORDOP_H__

#define SINGLE single
#define DASHED dashed
#define DASHSMALLGAP dashSmallGap
#include "fileop.h"
#include <fstream>
#include <vector>

class WordOp{
    public:
        WordOp(std::string filepath = "");

        void open(std::string filepath = "");

        void close();

        int replaceText(std::string mark, std::string replaced_text);

        int replaceImage(std::string mark, std::string replaced_image);

        int addInfoRecursive(int index, std::string text_file_path, std::string image_file_path);

        int createTable(int pos, int rows, int cols, int row_height, std::string border_line_type, std::string border_line_color, std::string text_color, std::string text);

        int createTable(int pos, int rows, int cols, std::vector< std::vector<std::string> > border_line_type, std::vector< std::vector<std::string> > border_line_colors, std::vector<std::string> text_color, std::vector<std::string> text);

    private:
        void readXml(std::string &xml_file, std::string filepath = "");
        void writeXml(std::string &xml_file, std::string filepath = "");
    private:
        std::string filepath;
        std::string cache_path;
        std::string document_xml, style_xml;

};

#endif
