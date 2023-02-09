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
        void static zipFolder(const std::string ziped_path);

        std::string static unzipFolder(const std::string unziped_path);

        void static dirIterate(std::string dir_path, void(*operation)(std::string, void *, int ), void *arg, int flag);

        void static deleteCache(std::string cache_path);

    private:

        void static dirZiper(std::string path, void *arg, int flag);
        void static dirRemover(std::string path, void *arg, int flag);

    private:

};

#endif
