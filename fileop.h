/*********************************************************************************************************************************************
File name:  fileop.h
Author: Li Xuan
Version: 1.0
Date:  2023-2-19
Description:  This module is used to zip and unzip files with suffix ".docx", it depends on zlib, zipper, minizips and Qt
Function List: 
    zipFolder:  Zip a folder
    unzipFolder:  Unzip a folder
    dirIterator:  Iterate a folder recursively. It also provide a function pointer argument that performs specific actions during the iteration
    deleteCache:  Delete the cache folder
    dirZipper:  A specific function that uses as a argument passses to the dirIterator function

***********************************************************************************************************************************************/

#ifndef __FILEOP_H__
#define __FILEOP_H__

#include "zipper.h"
#include "unzipper.h"
#include <string.h>
#include <string>
#include <assert.h>
#include <iostream>
#include <stdio.h>
#include <QFileInfo>
#include <QDirIterator>
#include <QFile>

#define MAX_PATH 1000


class FileOp{
    public:
        FileOp();
        ~FileOp();
        void static zipFolder(const std::string ziped_path);

        std::string static unzipFolder(const std::string unziped_path);

        void static dirIterate(std::string dir_path, void(*operation)(std::string, void *, int ), void *arg, int flag);

        void static deleteCache(std::string cache_path);

    private:

        void static dirZiper(std::string path, void *arg, int flag);
    private:

};

#endif
