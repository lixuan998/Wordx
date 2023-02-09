#include "wordop.h"

WordOp::WordOp(std::string filepath)
{
    cache_path = "";
    this->filepath = filepath;
    image_rels = "<Relationship Id=\"${rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"${target}\"/>";
    std::vector<std::string> str;
    //analyzeXml(str, "");
}

void WordOp::open(std::string filepath)
{
    if(!filepath.empty()) this->filepath = filepath;
    cache_path = FileOp::unzipFolder(this->filepath);
    //cache_path = "/home/climatex/Documents/wordx/loop_test.cache";
    std::cout << "open path: " << cache_path << std::endl;
    readXml(document_xml, (cache_path + "/word/document.xml"));
}

void WordOp::close()
{

    std::cout << "cache path: " << cache_path << std::endl;
    FileOp::zipFolder(cache_path);
    document_xml.clear();
    cache_path.clear();
    filepath.clear();
}
int WordOp::replaceText(std::string mark, std::string replaced_text, std::string &rep_str)
{
    if(rep_str == "")
    {
        if(document_xml.empty()) return -1;
        int p1 = myFind(document_xml, "<w:p>", 1);
        int p2 = myFind(document_xml, "</w:p>", -1);

        std::string tmp = document_xml.substr(p1, p2 + 6 - p1);

        std::vector<std::string> ana;

        analyzeXml(ana, tmp);

        for(int i = 0; i < ana.size(); ++ i)
        {
            std::cout << "ana i : " << i << " val: " << ana[i] << std::endl;
        }
        //std::cout << "tt: " << std::endl << tmp << std::endl;

        int pos = -1;
        while((pos = document_xml.find(mark, pos + 1)) != std::string::npos)
        {
            document_xml.replace(pos, mark.length(), replaced_text);
        }
    }
    else
    {
        if(rep_str.empty()) return -1;
        int p1 = myFind(rep_str, "<w:p>", 1);
        int p2 = myFind(rep_str, "</w:p>", -1);

        std::string tmp = rep_str.substr(p1, p2 + 6 - p1);

        std::vector<std::string> ana;

        analyzeXml(ana, tmp);

        for(int i = 0; i < ana.size(); ++ i)
        {
            std::cout << "ana i : " << i << " val: " << ana[i] << std::endl;
        }
        //std::cout << "tt: " << std::endl << tmp << std::endl;

        int pos = -1;
        while((pos = rep_str.find(mark, pos + 1)) != std::string::npos)
        {
            rep_str.replace(pos, mark.length(), replaced_text);
        }
    }
    return 0;
}

int WordOp::replaceImage(std::string mark, std::string replaced_image)
{
    std::fstream f_mark, f_image;
    int rels_index = 1;
    bool found = false;
    f_mark.open((cache_path + "/word/media/"), std::ios::in);
    if(!f_mark.is_open())
    {
        QDir tmp_dir;
        tmp_dir.mkpath(QString::fromStdString(cache_path + "/word/media/"));
    }
    else f_mark.close();
    std::string suffixs[5] = {".jpg", ".jpeg", ".bmp", ".webp", ".png"};
    for(int i = 0;i < 5; ++ i)
    {
        f_mark.open((cache_path + "/word/media/" + mark + suffixs[i]), std::ios::in);
        if(f_mark.is_open())
        {
            mark += suffixs[i];
            found = true;
            f_mark.close();
            break;
        }
    }
    if(found == false)
    {
        mark += ".png";
        readXml(doc_rels_xml, (cache_path + "/word/_rels/document.xml.rels"));
        int pos;
        while(true)
        {
            pos = doc_rels_xml.find(("rId" + std::to_string(rels_index)));
            if(pos == std::string::npos)
            {
                std::string new_image_rels = image_rels;
                pos = new_image_rels.find("${rId}");
                new_image_rels.replace(pos, 6, ("rId" + std::to_string(rels_index)));
                pos = new_image_rels.find("${target}");
                new_image_rels.replace(pos, 9, ("media/" + mark));

                pos = doc_rels_xml.find("</Relationships>");
                doc_rels_xml.replace(pos, 16, new_image_rels + "</Relationships>");

                writeXml(doc_rels_xml, (cache_path + "/word/_rels/document.xml.rels"));
                break;
            }

            ++ rels_index;
        }
    }
    f_mark.open((cache_path + "/word/media/" + mark), std::ios::out | std::ios::ate);
    std::cout << "mark path: " << (cache_path + "/word/media/" + mark) << "  mark open or not ? " << f_mark.is_open() << std::endl;
    f_mark.clear();
    f_image.open(replaced_image, std::ios::in);
    std::cout << "image open or not ? " << f_image.is_open() << std::endl;
    char buf[1024];
    memset(buf, 0, sizeof(buf));

    while(!f_image.eof())
    {
        f_image.read(buf, 1024);
        f_mark.write(buf, 1024);
    }

    f_mark.close();
    f_image.close();
    if(found == false) return rels_index;
    else return 0;
}

int WordOp::addInfoRecursive(int index, std::string info_file_path)
{
    QDir tempdir("/media");
    int image_sn = tempdir.count();
    int loop_start, loop_end, tmp_pos;
    std::vector<std::string> rep_model;
    std::vector<std::vector<std::string>> rep_info;
    std::string tmp, rep_res;

    readList(rep_info, info_file_path);

    int pos_start = myFind(document_xml, "${START_LOOP}", index);
    if(pos_start == std::string::npos)
    {
        std::cout << "can not find " << index << " th loop start mark, over range";
        return -2;
    }
    loop_start = myFind(document_xml.substr(0, pos_start), "<w:p>", -1);

    std::cout << "loop start: " << loop_start << std::endl << document_xml.substr(loop_start) << std::endl;

    int pos_end = myFind(document_xml, "${END_LOOP}", index);
    if(pos_end == std::string::npos)
    {
        std::cout << "can not find " << index << " th loop end mark, over range";
        return -3;
    }
    std::cout << "pos_end: " << pos_end << std::endl;
    loop_end = pos_end + myFind((document_xml.substr(pos_end, document_xml.size())), "</w:p>", 1);
    if(loop_end != std::string::npos) loop_end += 6;
    std::cout << "loop end: " << loop_end << std::endl << document_xml.substr(loop_start, loop_end - loop_start) << std::endl << "ENDEDNDNDND" << std::endl;
    tmp = document_xml.substr(pos_start + 13, pos_end - pos_start - 13);

    pos_start = myFind(tmp, "<w:p>", 1);
    pos_end = myFind(tmp, "<w:p>", -1);

    analyzeXml(rep_model, tmp.substr(pos_start, pos_end - pos_start));

    for(int i = 0; i < rep_model.size(); ++ i)
    {

        if(rep_model[i].find("${LOOP_IMAGE}") != std::string::npos)
        {
            rep_model[i].clear();
            readXml(rep_model[i], "../wordx/models/image.xml");
        }
        //std::cout << "i: " << i << " " << rep_model[i] << std::endl << std::endl << std::endl;
    }
    for(int i = 0; i < rep_info.size(); ++ i)
    {

        //std::cout << "rep model size: " << rep_model.size() << std::endl;
        int info_index = 0;
        for(int j = 0; j < rep_model.size(); ++ j)
        {
            tmp = rep_model[j];

            if(tmp.find("${SN}") != std::string::npos)
            {
                int rid = replaceImage(std::to_string((++ image_sn)), rep_info[i][info_index ++]);
                replaceText("${SN}", ("rId" + std::to_string(rid)), tmp);
                replaceText("${H}", "align", tmp);
                replaceText("${HARG}", "center", tmp);

                replaceText("${V}", "posOffset", tmp);
                replaceText("${VARG}", "653", tmp);
            }

            if(tmp.find("${LOOP_SN}") != std::string::npos)
            {
                replaceText("${LOOP_SN}", rep_info[i][info_index ++], tmp);
            }

            if(tmp.find("${LOOP_DEFECT}") != std::string::npos)
            {
                replaceText("${LOOP_DEFECT}", rep_info[i][info_index ++], tmp);
            }

            if(tmp.find("${LOOP_SURFACE}") != std::string::npos)
            {
                replaceText("${LOOP_SURFACE}", rep_info[i][info_index ++], tmp);
            }

            if(tmp.find("${LOOP_X}") != std::string::npos)
            {
                replaceText("${LOOP_X}", rep_info[i][info_index ++], tmp);
            }

            if(tmp.find("${LOOP_Y}") != std::string::npos)
            {
                replaceText("${LOOP_Y}", rep_info[i][info_index ++], tmp);
            }

            rep_res.append(tmp);

            //std::cout <<"j: " << j << " tmp: " << std::endl << tmp << std::endl;
        }

    }

    document_xml.replace(loop_start, loop_end - loop_start, rep_res);



    writeXml(document_xml, (cache_path + "/word/document.xml"));
    return 0;
}

int WordOp::createTable(int pos, int rows, int cols, int row_height, std::string border_line_type, std::string border_line_color, std::string text_color, std::string text)
{
    readXml(style_xml, cache_path + "/word/styles.xml");

    if(style_xml.find("Table Contents") == std::string::npos)
    {
        std::string table_style;
        readXml(table_style, "./xml_models/table_style.xml");


        //std::cout << std::endl;
        //std::cout << "last: " << last_pos << std::endl;

        //std::cout << std::endl;



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
        std::cout << "Error: Could not open file: " << filepath << std::endl;
        return;
    }
    while(!f.eof())
    {
        std::string line;
        getline(f, line);
        xml_file.append(line + "\n");
    }
    f.close();

    //std::cout << "read: " << xml_file << std::endl;
}

void WordOp::readList(std::vector<std::vector<std::string>> &str_list, std::string filepath)
{
    std::cout << "filepath: " << filepath << std::endl;
    std::fstream f(filepath, std::ios::in);
    if(!f.is_open())
    {
        std::cout << "fail to open file" << std::endl;
        return;
    }

    std::vector<std::string> tmp_vec;
    while(!f.eof())
    {
        std::string tmp;
        getline(f, tmp);
        if(!tmp.empty()) tmp_vec.push_back(tmp);
        else
        {
            if(tmp_vec.size() != 1) str_list.push_back(tmp_vec);
            tmp_vec.clear();
        }
    }
    f.close();

//    std::cout << "str_list size: " << str_list.size() << " child size : " << str_list[0].size() << std::endl;
//    for(int i = 0; i < str_list.size(); ++ i)
//    {
//        for(int j = 0; j < str_list[i].size(); ++ j)
//        {
//            std::cout << str_list[i][j] << std::endl;
//        }
//    }
    return;
}

int WordOp::myFind(std::string src, std::string des, int index)
{
    //std::cout << "myFind src: " << std::endl << src << std::endl;
    int _index = 1;
    int pos = 0;
    pos = src.find(des);
    if(index == -1)
    {
        while(true)
        {
            int tmp_pos = src.find(des, pos + 1);
            if(tmp_pos != std::string::npos) pos = tmp_pos;
            else return pos;
        }
    }
    while(_index != index)
    {
        pos = src.find(des, pos + 1);
        ++ _index;
        if(pos == std::string::npos) break;
    }
    return pos;
}

void WordOp::analyzeXml(std::vector<std::string> &analysis_xml, std::string origin_xml)
{
    int pos_start, pos_end, index = 1;
    pos_start = myFind(origin_xml, "<w:p>", index);
    pos_end = myFind(origin_xml, "</w:p>", index);
    while(pos_start != std::string::npos)
    {
        analysis_xml.push_back(origin_xml.substr(pos_start, pos_end + 6 - pos_start));
        ++index;
        pos_start = myFind(origin_xml, "<w:p>", index);
        pos_end = myFind(origin_xml, "</w:p>", index);
    }
}

void WordOp::writeXml(std::string &xml_file, std::string filepath)
{
    std::fstream f;
    f.open(filepath, std::ios::out);
    f << xml_file << std::endl;
    f.close();

    //std::cout << xml_file << std::endl;
}

std::vector<int> WordOp::findImgModelPos(std::string src, int index)
{
    std::vector<int> ret;
    ret.push_back(-1);
    ret.push_back(-1);

    int pos_mid = myFind(src, "${LOOP_IMAGE}", index);
    if(pos_mid == std::string::npos)
    {
        std::cout << "can not find " << index << " th image model" << std::endl;
        return ret;
    }
    std::string front = src.substr(0, pos_mid);
    std::string back = src.substr(pos_mid);
    ret[0] = myFind(front, "<w:p>", -1);
    ret[1] = myFind(back, "</w:p>", 1);

    std::string res = src.substr(ret[0], ret[1] + 6 - ret[0]);
    std::cout << "res: " << std::endl << res << std::endl;
    return ret;
}
