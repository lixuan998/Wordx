#include <iostream>
#include <string>
#include <cstdio>

#include "fileop.h"
#include "wordop.h"

int main(int argc, char **argv)
{
    if(argc < 2)
    {
        std::cerr << "Usage: ./word <word_file_path>\n";
        return -1;
    }
    WordOp op(argv[1]);
    op.open();

    op.replaceImageFromFile(":/new/demo_files/demo_files/mark_list", ":/new/demo_files/demo_files/img_list");
    op.replaceText(":/new/demo_files/demo_files/mark_info", ":/new/demo_files/demo_files/replace_info");
    std::vector<std::string> table_files, recur_files;
    std::vector<int> table_indexes, recur_indexes;

    table_files.push_back(":/new/demo_files/demo_files/table_info");
    table_indexes.push_back(1);

    recur_files.push_back(":/new/demo_files/demo_files/first_rec_info_list");
    recur_files.push_back(":/new/demo_files/demo_files/second_rec_info_list");
    recur_indexes.push_back(1);
    recur_indexes.push_back(2);
    
    op.addTableRows(table_indexes, table_files);
    op.addInfoRecursive(recur_indexes, recur_files);
    
    op.close();

}
