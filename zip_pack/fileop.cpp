#include "fileop.h"

FileOp::FileOp()
{

}

void FileOp::dirIterate(std::string dir_path, void(*operation)(std::string path, void *arg0, int flag), void *arg, int flag)
{
   QDirIterator it(QString::fromStdString(dir_path));

   while(it.hasNext())
   {
      it.next();
      if(it.fileName() == "." || it.fileName() == "..") continue;
      if(it.fileInfo().isDir())
      {
         dirIterate(it.filePath().toStdString(), operation, arg, (flag + 1));
      }

      (operation)(it.filePath().toStdString(), arg, flag);
   }
}

void FileOp::zipFolder(const std::string unziped_path)
{

   std::string zip_path = unziped_path.substr(0, unziped_path.find_last_of('.')) + "_copy.docx";
   remove(zip_path.c_str());
   zipper::Zipper *zipper = new zipper::Zipper(zip_path, zipper::Zipper::openFlags::Overwrite);
   dirIterate(unziped_path.c_str(), dirZiper, (void *) zipper, 0);
   zipper->close();
   delete zipper;
   zipper = nullptr;
   deleteCache(unziped_path.c_str());

}

std::string FileOp::unzipFolder(const std::string ziped_path)
{
   int dot_pos = ziped_path.find_last_of('.');
   std::string dir_path = ziped_path.substr(0, dot_pos) + ".cache";
   zipper::Unzipper *unzipper = new zipper::Unzipper(ziped_path);
   unzipper->extract(dir_path);
   return dir_path;
}

void FileOp::deleteCache(std::string cache_path)
{
    dirIterate(cache_path, dirRemover, NULL, 0);
    QFile::remove(QString::fromStdString(cache_path));
}

void FileOp::dirZiper(std::string path, void *arg, int flag)
{
   int cnt = 0;
   if(flag != 0) return;
   zipper::Zipper *zipper = reinterpret_cast<zipper::Zipper *>(arg);
   zipper->add(path);
}

void FileOp::dirRemover(std::string path, void *arg, int flag)
{
    QFile::remove(QString::fromStdString(path));
}
