#include "fileop.h"

FileOp::FileOp()
{

}

void FileOp::dirIterate(const char *dir_path, void(*operation)(const char *path, void *arg0), void *arg)
{
    DIR *dir;
    struct dirent *entry;
    dir = opendir(dir_path);
    if(dir == NULL) return;
    assert(dir);

    char fullpath[MAX_PATH];
    memset(fullpath, 0, sizeof(fullpath));
    struct stat s;

    while ((entry = readdir(dir)))
    {
      if (!strcmp(entry->d_name, ".\0") || !strcmp(entry->d_name, "..\0")) continue;
      snprintf(fullpath, sizeof(fullpath), "%s/%s", dir_path, entry->d_name);
      stat(fullpath, &s);
      if (S_ISDIR(s.st_mode)) dirIterate(fullpath, operation, arg);

      (operation)(fullpath, arg);
    }

    closedir(dir);
}
void FileOp::zipFolder(const std::string unziped_path)
{

    std::string zip_path = unziped_path.substr(0, unziped_path.find_last_of('.')) + "_copy.docx";
    remove(zip_path.c_str());
    zipper::Zipper *zipper = new zipper::Zipper(zip_path, zipper::Zipper::openFlags::Overwrite);
    dirIterate(unziped_path.c_str(), dirZiper, (void *) zipper);
    zipper->close();
    delete zipper;
    zipper = nullptr;
    deleteCache(unziped_path.c_str());

}

std::string FileOp::unzipFolder(const std::string ziped_path)
{
    int dot_pos = ziped_path.find_last_of('.');
    std::string dir_path = ziped_path.substr(0, dot_pos) + ".cache";
    zip_extract(ziped_path.c_str(), dir_path.c_str(), NULL, NULL);
    return dir_path;
}

void FileOp::deleteCache(const char * cache_path)
{
    dirIterate(cache_path, dirRemover, NULL);
    remove(cache_path);
}

void FileOp::dirZiper(const char *path, void *arg)
{
    int cnt = 0;
    for(int i = 0; i < strlen(path); ++ i)
    {
      if(path[i] == '/' || path[i] == '\\') cnt ++;
      if(cnt > 1) return;
    }
    std::cout << "zip file " << path << std::endl;
    zipper::Zipper *zipper = reinterpret_cast<zipper::Zipper *>(arg);
    zipper->add(path);
}

void FileOp::dirRemover(const char *path, void *arg)
{
    remove(path);
}
