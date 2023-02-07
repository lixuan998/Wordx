#include <iostream>
#include <string>
#include <cstdio>

#include "fileop.h"
#include "wordop.h"

int main(int argc, char **argv)
{
    WordOp op("test2.docx");
    op.open();
    std::string tmp;
    //op.createTable(0, 0, 0, "", "", "", "");
    op.replaceText("${1}", "替换第一个标签");
    op.replaceText("${2}", "替换第二个标签");
    op.replaceImage("image1", "/home/climatex/Pictures/Screenshots/A.jpeg");
    op.replaceImage("image2", "/home/climatex/Pictures/Screenshots/B.png");

    op.replaceImage("image3", "/home/climatex/Pictures/Screenshots/C.png");

     op.close();
    // FileOp op;
    // if(strcmp(argv[1], "zip") == 0)
    // {
    //     op.zipFolder(argv[2]);
    // }
    // else if(strcmp(argv[1], "unzip") == 0)
    // {
    //     std::cout << "unzip" << std::endl;
    //     op.unzipFolder(argv[2]);
    // }
    //op.unzipFolder(filepath);

}
