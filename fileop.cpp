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
   QFile::remove(QString::fromStdString(zip_path));
   zipper::Zipper *zipper = new zipper::Zipper(zip_path, zipper::Zipper::openFlags::Overwrite);
   dirIterate(unziped_path, dirZiper, (void *) zipper, 0);
   zipper->close();
   delete zipper;
   zipper = nullptr;
   deleteCache(unziped_path);

}

std::string FileOp::unzipFolder(const std::string ziped_path)
{
   int dot_pos = ziped_path.find_last_of('.');
   std::string dir_path = ziped_path.substr(0, dot_pos) + ".cache";
   zipper::Unzipper *unzipper = new zipper::Unzipper(ziped_path);
   unzipper->extract(dir_path);
   std::cout << "unzippe folder " << dir_path << std::endl;
   return dir_path;
}

void FileOp::deleteCache(std::string cache_path)
{
    QDir tmp_dir(QString::fromStdString(cache_path));
    tmp_dir.removeRecursively();
}

void FileOp::dirZiper(std::string path, void *arg, int flag)
{
   if(flag != 0) return;

   zipper::Zipper *zipper = reinterpret_cast<zipper::Zipper *>(arg);
   zipper->add(path);
   std::cout << "zip file: " << path << std::endl;
}

