#include <iostream>
#include <string>
#include <cstdio>
#include <typeinfo>
#include <opencv2/core/core.hpp>

#include "fileop.h"
#include "wordop.h"

int main(int argc, char **argv)
{
    std::string word_path;
    if(argc < 2)
    {
        word_path = "standard_model.docx";
    }
    else word_path = argv[1];
    Info info;
    std::string defects[]{"污渍", "划痕", "氧化", "橡胶", "苹果"};
    std::string surfaces[]{"上表面", "下表面", "上表面", "下表面", "上表面"};
    std::string xs[]{"10", "20", "30", "40", "50"};
    std::string ys[]{"1000", "2000", "3000", "4000", "5000"};
    std::string sns[]{"1", "2", "3", "4", "5"};

    std::string arg1[]{"1", "2", "3", "4", "5"};
    std::string arg2[]{"污渍", "划痕", "氧化", "橡胶", "苹果"};
    std::string arg3[]{"1000", "2000", "3000", "4000", "5000"};

    std::vector<Info> rec_infos, table_infos;
    for(int i = 0; i < 5; ++ i)
    {
        info.clear();
        info.addInfo("${LP_SN}", sns[i]);
        info.addInfo("${LP_DEFECT}", defects[i]);
        info.addInfo("${LP_SURFACE}", surfaces[i]);
        info.addInfo("${LP_X}", xs[i]);
        info.addInfo("${LP_Y}", ys[i]);
        info.addInfo("${LP_IMAGE}", cv::imread("../demo_files/steel1.jpeg"));
        rec_infos.push_back(info);
    }

    for(int i = 0; i < 5; ++ i)
    {
        info.clear();
        info.addInfo("${TB_ARG1}", arg1[i]);
        info.addInfo("${TB_ARG2}", arg2[i]);
        info.addInfo("${TB_ARG3}", arg3[i]);
        info.addInfo("${TB_ARG4}", cv::imread("../demo_files/steel1.jpeg"));
        table_infos.push_back(info);
    }

    std::vector<std::string> marks{"${ST_UP_SURFACE}"};
    std::vector<std::string> replace_image_paths{"../demo_files/steel1.jpeg"};
    std::vector<cv::Mat> replace_mat_images{cv::imread("../demo_files/steel1.jpeg")};
    WordOp op(word_path);
    op.open();
    op.replaceImageFromMat(marks, replace_mat_images);
    //op.replaceImageFromFile(":/new/demo_files/demo_files/mark_list", ":/new/demo_files/demo_files/img_list");
    op.replaceText(":/new/demo_files/demo_files/mark_info", ":/new/demo_files/demo_files/replace_info");
    std::vector<std::string> table_files, recur_files;
    std::vector<int> table_indexes, recur_indexes;

    table_files.push_back(":/new/demo_files/demo_files/table_info");
    table_indexes.push_back(1);

    recur_files.push_back(":/new/demo_files/demo_files/first_rec_info_list");
    recur_files.push_back(":/new/demo_files/demo_files/second_rec_info_list");
    recur_indexes.push_back(1);
    recur_indexes.push_back(2);
    
    op.addTableRows(table_indexes, table_infos);
    op.addInfoRecursive(recur_indexes, rec_infos);
    
    op.close();

    std::string gen_path = word_path.substr(0, word_path.find_last_of(".")) + "_copy.docx";
    std::cout << "\033[32m" << "----------New file \"" << gen_path << "\" generated----------" << std::endl;}
