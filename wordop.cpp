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

    std::cout << "open path: " << cache_path << std::endl;
    readXml(document_xml, (cache_path + "/word/document.xml"));
}

void WordOp::close()
{

    writeXml(document_xml, (cache_path + "/word/document.xml"));
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

        analyzeXml(ana, tmp, "w:p");

        for(int i = 0; i < ana.size(); ++ i)
        {
            std::cout << "ana i : " << i << " val: " << ana[i] << std::endl;
        }
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

        analyzeXml(ana, tmp, "w:p");

        int pos = -1;
        while((pos = rep_str.find(mark, pos + 1)) != std::string::npos)
        {
            rep_str.replace(pos, mark.length(), replaced_text);
        }
    }
    return 0;
}
int WordOp::replaceText(std::vector<std::string> marks, std::vector<std::string> replaced_texts, std::string &rep_str)
{
    if(marks.size() != replaced_texts.size())
    {
        std::cout << "Mark size does not match ReplaceText size" << std::endl;
        return -1;
    }
    for(int i = 0; i < marks.size(); ++ i)
    {
        replaceText(marks[i], replaced_texts[i], rep_str);
    }
    return 0;
}
int WordOp::replaceText(std::string mark_file, std::string replaced_text_file)
{
    std::vector<std::vector<std::string>> rep_info, mark_info;
    std::vector<std::string> marks, reps;
    readList(mark_info, mark_file);
    readList(rep_info, replaced_text_file);
    for(int i = 0; i < mark_info.size(); ++ i)
    {
        for(int j = 0; j < mark_info[i].size(); ++ j)
        {
            marks.push_back(mark_info[i][j]);
            reps.push_back(rep_info[i][j]);
        }
    }
    if(marks.size() != reps.size())
    {
        std::cout << "Items in mark file does not equal to items in replaced_text file" << std::endl;
        return -1;
    }
    std::string emp = "";
    replaceText(marks, reps, emp);
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
    f_mark.clear();
    f_image.open(replaced_image, std::ios::in);
    if(!f_image.is_open())
    {
        std::cout << "Error can not open image " << replaced_image << std::endl;
        return -1;
    }
    std::cout << replaced_image << " image open or not ? " << f_image.is_open() << std::endl;
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

int WordOp::addInfoRecursive(std::vector<int> indexs, std::vector<std::string> info_file_path)
{
    QDir tempdir(QString::fromStdString(cache_path + "/word/media"));
    int image_sn = tempdir.count() - 2;
    std::cout << "image_sn: " << image_sn << std::endl;
    
    std::string tmp;
    for(int dex = indexs.size() - 1; dex >= 0; dex --)
    {
        int loop_start, loop_end, tmp_pos, index;
        std::vector<std::string> rep_model;
        std::vector<std::vector<std::string>> rep_info;
        std::string rep_res;

        loop_start = loop_end = tmp_pos = index = 0;
        rep_model.clear();
        rep_info.clear();
        rep_res.clear();
        index = indexs[dex];
        readList(rep_info, info_file_path[dex]);
        std::cout << "rep_info: " << std::endl;
        for(int j = 0; j < rep_info.size(); ++ j)
        {
            for(int k = 0; k < rep_info[j].size(); ++ k)
            {
                std::cout << rep_info[j][k] << std::endl;
            }
            std::cout << std::endl;
        }


        int pos_start = myFind(document_xml, "${LP_START}", index);
        if(pos_start == std::string::npos)
        {
            std::cout << "can not find " << index << " th loop start mark, over range";
            return -2;
        }
        loop_start = findAround(document_xml, "<w:p>", pos_start, UP);
        if(loop_start == std::string::npos || loop_start == -1) return -3;
        int pos_end = myFind(document_xml, "${LP_END}", index);
        if(pos_end == std::string::npos)
        {
            std::cout << "can not find " << index << " th loop end mark, over range" << std::endl;
            return -3;
        }
        loop_end = pos_end + findAround(document_xml, "</w:p>", pos_end, DOWN);
        if(loop_end == std::string::npos || loop_end == -1) return -3;
        loop_end += 6;
        tmp = document_xml.substr(pos_start, pos_end - pos_start);
        pos_start = myFind(tmp, "<w:p>", 1);
        pos_end = myFind(tmp, "<w:p>", -1);

        analyzeXml(rep_model, tmp.substr(pos_start, pos_end - pos_start), "w:p");

        for(int i = 0; i < rep_model.size(); ++ i)
        {

            if(rep_model[i].find("${LP_IMAGE}") != std::string::npos)
            {
                rep_model[i].clear();
                readXml(rep_model[i], "../xml_models/image.xml");
            }
        }

        for(int i = 0; i < rep_info.size(); ++ i)
        {

            int info_index = 0;
            for(int j = 0; j < rep_model.size(); ++ j)
            {
                tmp = rep_model[j];
                tmp_pos = tmp.find("${IMAGE_SN}");
                if(tmp_pos != std::string::npos)
                {
                    int rid = replaceImage(("insert_image" + std::to_string((++ image_sn))), rep_info[i][info_index ++]);
                    replaceText("${IMAGE_SN}", ("rId" + std::to_string(rid)), tmp);
                    replaceText("${H}", "align", tmp);
                    replaceText("${HARG}", "center", tmp);
                    replaceText("${V}", "posOffset", tmp);
                    replaceText("${VARG}", "653", tmp);
                }
                else
                {
                    tmp_pos = tmp.find("${LP_");
                    if(tmp_pos != std::string::npos)
                    {
                        std::string tmp_mark;
                        for(int k = tmp_pos; ; ++ k)
                        {
                            tmp_mark.push_back(tmp[k]);
                            if(tmp[k] == '}') break;
                        }

                        //std::cout <<"tmp_mark: " <<  tmp_mark << " replace with: " << rep_info[i][info_index] << std::endl;
                        replaceText(tmp_mark, rep_info[i][info_index ++], tmp);
                    }
                }
                rep_res.append(tmp);
            }
        }

        document_xml.replace(loop_start, loop_end - loop_start, rep_res);
        std::cout << "REPLACEMENT" << std::endl;
    }
    
    writeXml(document_xml, (cache_path + "/word/document.xml"));
    return 0;
}

int WordOp::addInfoRecursive(std::vector<int> indexs, std::vector<std::vector<std::vector<std::string>>> rep_infos)
{
    QDir tempdir(QString::fromStdString(cache_path + "/word/media"));
    int image_sn = tempdir.count() - 2;
    std::cout << "image_sn: " << image_sn << std::endl;
    
    std::string tmp;
    for(int i = indexs.size() - 1; i >= 0; -- i)
    {
        int loop_start, loop_end, tmp_pos, index;
        std::vector<std::string> rep_model;
        std::vector<std::vector<std::string>> rep_info;
        std::string rep_res;

        loop_start = loop_end = tmp_pos = index = 0;
        rep_model.clear();
        rep_info.clear();
        rep_res.clear();
        index = indexs[i];
        rep_info = rep_infos[i];
        if(rep_info.empty())
        {
            std::cout << "Error rep_info empty" << std::endl;
            return -1;
        }
        int pos_start = myFind(document_xml, "${LP_START}", index);
        if(pos_start == std::string::npos)
        {
            std::cout << "can not find " << index << " th loop start mark, over range";
            return -2;
        }
        loop_start = findAround(document_xml, "<w:p>", pos_start, UP);
        if(loop_start == std::string::npos || loop_start == -1) return -3;
        int pos_end = myFind(document_xml, "${LP_END}", index);
        if(pos_end == std::string::npos)
        {
            std::cout << "can not find " << index << " th loop end mark, over range" << std::endl;
            return -3;
        }
        loop_end = findAround(document_xml, "</w:p>", pos_end, DOWN);
        if(loop_end == std::string::npos || loop_end == -1) return -3;
        loop_end += 6 + pos_end;
        tmp = document_xml.substr(pos_start, pos_end - pos_start );
        pos_start = myFind(tmp, "<w:p>", 1);
        pos_end = myFind(tmp, "<w:p>", -1);

        analyzeXml(rep_model, tmp.substr(pos_start, pos_end - pos_start), "w:p");

        for(int i = 0; i < rep_model.size(); ++ i)
        {

            if(rep_model[i].find("${LP_IMAGE}") != std::string::npos)
            {
                rep_model[i].clear();
                readXml(rep_model[i], "../xml_models/image.xml");
            }
        }

        for(int i = 0; i < rep_info.size(); ++ i)
        {

            int info_index = 0;
            for(int j = 0; j < rep_model.size(); ++ j)
            {
                tmp = rep_model[j];
                tmp_pos = tmp.find("${IMAGE_SN}");
                if(tmp_pos != std::string::npos)
                {
                    int rid = replaceImage(("insert_image" + std::to_string((++ image_sn))), rep_info[i][info_index ++]);
                    replaceText("${IMAGE_SN}", ("rId" + std::to_string(rid)), tmp);
                    replaceText("${H}", "align", tmp);
                    replaceText("${HARG}", "center", tmp);
                    replaceText("${V}", "posOffset", tmp);
                    replaceText("${VARG}", "653", tmp);
                }
                else
                {
                    tmp_pos = tmp.find("${LP_");
                    if(tmp_pos != std::string::npos)
                    {
                        std::string tmp_mark;
                        for(int k = tmp_pos; ; ++ k)
                        {
                            tmp_mark.push_back(tmp[k]);
                            if(tmp[k] == '}') break;
                        }

                        //std::cout <<"tmp_mark: " <<  tmp_mark << " replace with: " << rep_info[i][info_index] << std::endl;
                        replaceText(tmp_mark, rep_info[i][info_index ++], tmp);
                    }
                }
                rep_res.append(tmp);
            }
        }

        document_xml.replace(loop_start, loop_end - loop_start, rep_res);

        std::cout << "replace" << std::endl;
    }

    return 0;
}

int WordOp::addTableRows(std::vector<int> indexs, std::vector<std::string> info_file_path)
{
    std::string tmp;
    for(int dex = indexs.size() - 1; dex >= 0; -- dex)
    {
        int loop_start, loop_end, tmp_pos, index;
        std::vector<std::string> rep_model;
        std::vector<std::vector<std::string>> rep_info;
        std::string rep_res;

        loop_start = loop_end = tmp_pos = index = 0;
        rep_model.clear();
        rep_info.clear();
        rep_res.clear();
        index = indexs[dex];
        readList(rep_info, info_file_path[dex]);

        int pos_start = myFind(document_xml, "${TB_START}", index);
        if(pos_start == std::string::npos)
        {
            std::cout << "The " << index << "th TB_START not found" << std::endl;
            return -1;
        }
        loop_start = findAround(document_xml, "<w:tr>", pos_start, UP);
        if(loop_start == std::string::npos || loop_start == -1)
        {
            std::cout << "The " << index << "th loop_start not found" << std::endl;
            return -2;
        }

        int pos_end = myFind(document_xml, "${TB_END}", index);
        if(pos_end == std::string::npos)
        {
            std::cout << "The " << index << "th TB_END not found" << std::endl;
            return -1;
        }
        loop_end = findAround(document_xml, "</w:tr>", pos_end, DOWN);
        if(loop_end == std::string::npos || loop_end == -1)
        {
            std::cout << "The " << index << "th loop_end not found" << std::endl;
            return -2;
        }
        loop_end += 7 + pos_end;
        tmp = document_xml.substr(pos_start, pos_end - pos_start);
        pos_start = myFind(tmp, "<w:tr>", 1);
        pos_end = myFind(tmp, "<w:tr>", -1);

        analyzeXml(rep_model, tmp.substr(pos_start, pos_end - pos_start), "w:tr");

        // for(int i = 0; i < rep_model.size(); i++)
        // {
        //     std::cout << "i: " << i << std::endl;
        //     std::cout << rep_model[i] << std::endl;
        // }

        for(int i = 0; i < rep_info.size(); ++ i)
        {
            for(int j = 0; j < rep_info[i].size(); ++ j)
            {
                std::string rep_mark = "${TB_ARG" + std::to_string(j + 1) + "}";
                std::cout << "rep mark: " << rep_mark << std::endl;
                for(int k = 0; k < rep_model.size(); ++ k)
                {
                    if(rep_model[k].find(rep_mark) != std::string::npos)
                    {
                        replaceText(rep_mark, rep_info[i][j], rep_model[k]);
                    }
                }
            }
            for(int j = 0; j < rep_model.size(); ++ j) rep_res.append(rep_model[j]);
        }
        // std::cout << "part to replace: " << std::endl;

        // std::cout << document_xml.substr(loop_start, loop_end - loop_start) << std::endl;
        document_xml.replace(loop_start, loop_end - loop_start, rep_res);
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
            for(auto a : tmp_vec)
            {
                if(!a.empty())
                {
                    str_list.push_back(tmp_vec);
                    break;
                }
            }
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

int WordOp::findAround(std::string src, std::string des, int cur_pos, int direction)
{
    std::string tmp;
    int ret;
    if(direction == UP)
    {
        tmp = src.substr(0, cur_pos);
        ret = myFind(tmp, des, -1);
    }
    else if(direction == DOWN)
    {
        tmp = src.substr(cur_pos);
        ret = myFind(tmp, des, 1);
    }
    else
    {
        ret = -1;
        std::cout << "findAround argument error" << std::endl;
    }

    return ret;
}

void WordOp::analyzeXml(std::vector<std::string> &analysis_xml, std::string origin_xml, std::string flag)
{
    analysis_xml.clear();
    int pos_start, pos_end, index = 1;
    pos_start = myFind(origin_xml, ("<" + flag + ">"), index);
    pos_end = myFind(origin_xml, ("</" + flag + ">"), index);
    while(pos_start != std::string::npos)
    {
        analysis_xml.push_back(origin_xml.substr(pos_start, pos_end + ("</" + flag + ">").size() - pos_start));
        ++index;
        pos_start = myFind(origin_xml, ("<" + flag + ">"), index);
        pos_end = myFind(origin_xml, ("</" + flag + ">"), index);
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

