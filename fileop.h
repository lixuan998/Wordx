#ifndef __FILEOP_H__
#define __FILEOP_H__

#include <zip.h>
#include <Zipper.hpp>
#include <string.h>
#include <string>
#include <dirent.h>
#include <sys/stat.h>
#include <assert.h>
#include <iostream>
#include <stdio.h>

#define MAX_PATH 1000


class FileOp{
    public:
        FileOp();
        void static zipFolder(const std::string ziped_path);

        std::string static unzipFolder(const std::string unziped_path);

        void static dirIterate(const char *dir_path, void(*operation)(const char *, void *), void *arg);

        void static deleteCache(const char * cache_path);

    private:

        void static dirZiper(const char *path, void *arg);
        void static dirRemover(const char *path, void *arg);

    private:

};

#endif
