#include <iostream>
#include <string>
#include <cstdio>

#include "fileop.h"
#include "wordop.h"

int main(int argc, char **argv)
{
    WordOp op("/home/climatex/Documents/wordx/standard_model.docx");
    op.open();


    op.replaceText("/home/climatex/Documents/wordx/mark_info", "/home/climatex/Documents/wordx/replace_text");
    std::vector<std::string> table_files, recur_files;
    std::vector<int> table_indexes, recur_indexes;

    table_files.push_back("/home/climatex/Documents/wordx/table_info");
    table_indexes.push_back(1);

    recur_files.push_back("/home/climatex/Documents/wordx/info_list");
    recur_files.push_back("/home/climatex/Documents/wordx/info_list2");
    recur_indexes.push_back(1);
    recur_indexes.push_back(2);
    
    op.addTableRows(table_indexes, table_files);
    op.addInfoRecursive(recur_indexes, recur_files);
    
     op.close();

}
