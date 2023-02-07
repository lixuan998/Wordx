#include "wordop.h"

WordOp::WordOp(std::string filepath)
{
    cache_path = "";
    this->filepath = filepath;
}

void WordOp::open(std::string filepath)
{
    if(!filepath.empty()) this->filepath = filepath;
    cache_path = FileOp::unzipFolder(this->filepath);
    std::cout << "cc path: " << cache_path << std::endl;
}

void WordOp::close()
{
    std::cout << "cache path: " << cache_path << std::endl;
    FileOp::zipFolder(cache_path);
    document_xml.clear();
    cache_path.clear();
    filepath.clear();
}
int WordOp::replaceText(std::string mark, std::string replaced_text)
{
    readXml(document_xml, cache_path + "/word/document.xml");
    if(document_xml.empty()) return -1;
    int pos = 0;
    while((pos = document_xml.find(mark, pos)) != std::string::npos)
    {
        document_xml.replace(pos, mark.length(), replaced_text);
    }

    //std::cout << "after replace text: " << document_xml << std::endl;
    std::fstream f(cache_path + "/word/document.xml", std::ios::out);
    f << document_xml;
    f.close();
    return 0;
}

int WordOp::replaceImage(std::string mark, std::string replaced_image)
{
    std::fstream f_mark, f_image;
    f_mark.open(cache_path + "/word/media/", std::ios::in);
    if(!f_mark.is_open())
    {
        std::cout << "fail to open media file " << std::endl;
        return -1;
    }
    f_mark.close();
    std::string suffixs[5] = {".png", ".jpg", ".jpeg", ".bmp", ".webp"};
    for(int i = 0;i < suffixs->size(); ++ i)
    {
        f_mark.open(cache_path + "/word/media/" + mark + suffixs[i], std::ios::in);
        if(f_mark.is_open())
        {
            mark += suffixs[i];
            f_mark.close();
            break;
        }
    }
    f_mark.open(cache_path + "/word/media/" + mark, std::ios::out);
    f_mark.clear();
    f_image.open(replaced_image, std::ios::in);

    char buf[1024];
    memset(buf, 0, sizeof(buf));

    while(!f_image.eof())
    {
        f_image.read(buf, 1024);
        f_mark.write(buf, 1024);
    }

    f_mark.close();
    f_image.close();

    return 0;
}

int WordOp::addInfoRecursive(int index, std::string text_file_path, std::string image_file_path)
{

    return 0;
}

int WordOp::createTable(int pos, int rows, int cols, int row_height, std::string border_line_type, std::string border_line_color, std::string text_color, std::string text)
{
    readXml(style_xml, cache_path + "/word/styles.xml");

    if(style_xml.find("Table Contents") == std::string::npos)
    {
        std::string table_style;
        readXml(table_style, "./xml_models/table_style.xml");
        int last_pos = style_xml.find_last_of("</w:styles>");

        //std::cout << std::endl;
        //std::cout << "last: " << last_pos << std::endl;

        //std::cout << std::endl;
        style_xml.insert(last_pos - 10, table_style);


        writeXml(style_xml, cache_path + "/word/styles.xml");
    }

    return 0;
}

void WordOp::readXml(std::string &xml_file, std::string filepath)
{
    xml_file.clear();
    char buf[4086];
    std::fstream f(filepath, std::ios::in);
    if(!f.is_open())
    {
        std::cout << "Error: Could not open file" << std::endl;
        return;
    }
    while(!f.eof())
    {
        std::string line;
        getline(f, line);
        xml_file.append(line);
    }
    f.close();
    //std::cout << xml_file << std::endl;
}

void WordOp::writeXml(std::string &xml_file, std::string filepath)
{
    std::fstream f;
    f.open(filepath, std::ios::out);
    f << xml_file << std::endl;
    f.close();

    std::cout << xml_file << std::endl;
}
