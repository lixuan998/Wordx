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
        std::cerr << "Mark size does not match ReplaceText size" << std::endl;
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
        std::cerr << "Items in mark file does not equal to items in replaced_text file" << std::endl;
        return -1;
    }
    std::string emp = "";
    replaceText(marks, reps, emp);
    return 0;
}

int WordOp::replaceImage(std::string mark, std::string replace_image)
{
    int mark_pos = myFind(document_xml, mark, 1);
    QDir dir(QString::fromStdString(cache_path + "/word/media"));
    int image_sn = dir.count() - 1;
    int rid = addImage(("replace_image" + std::to_string(image_sn)), replace_image);
    std::string img_model;
    readXml(img_model, ":/new/xml_models/xml_models/image.xml");
        
    replaceText("${IMAGE_SN}", ("rId" + std::to_string(rid)), img_model);
    replaceText("${H}", "align", img_model);
    replaceText("${HARG}", "center", img_model);
    replaceText("${V}", "posOffset", img_model);
    replaceText("${VARG}", "653", img_model);

    while(mark_pos != std::string::npos)
    {
        int rep_start = findAround(document_xml, "<w:p>", mark_pos, UP);
        if(rep_start == std::string::npos || rep_start == -1)
        {
            std::cerr << "Error fail to find <w:p> around pos " << mark_pos << " using UP" << std::endl;
            return -1;
        }
        int rep_end = findAround(document_xml, "</w:p>", mark_pos, DOWN);
        if(rep_end == std::string::npos || rep_end == -1)
        {
            std::cerr << "Error fail to find </w:p> around pos " << mark_pos << " using DOWN" << std::endl;
            return -1;
        }
        rep_end += mark_pos + 6;
        document_xml.replace(rep_start, rep_end - rep_start, img_model);

        mark_pos = myFind(document_xml, mark, 1);
    }
    return 0;
}

int WordOp::replaceImageFromFile(std::string mark_file_path, std::string replace_image_file_path)
{
    std::vector<std::vector<std::string>> mark_file, replace_image_file;
    std::vector<std::string> marks, rep_images;
    readList(mark_file, mark_file_path);
    readList(replace_image_file, replace_image_file_path);
    for(int i = 0; i < mark_file.size(); ++ i)
    {
        for(int j = 0; j < mark_file[i].size(); ++ j)
        {
            marks.push_back(mark_file[i][j]);
        }
    }
    for(int i = 0; i < replace_image_file.size(); ++ i)
    {
        for(int j = 0; j < replace_image_file[i].size(); ++ j)
        {
            rep_images.push_back(replace_image_file[i][j]);
        }
    }

    if(rep_images.size() != marks.size())
    {
        std::cerr << "Error from replaceImageFromFile: rep_images size does not match marks" << std::endl;
        return -1;
    }

    for(int i = 0; i < marks.size(); ++ i)
    {
        replaceImage(marks[i], rep_images[i]);
    }

    return 0;
}

int WordOp::addInfoRecursive(std::vector<int> indexs, std::vector<std::string> info_file_path)
{
    QDir tempdir(QString::fromStdString(cache_path + "/word/media"));
    int image_sn = tempdir.count() - 2;    
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

        int pos_start = myFind(document_xml, "${LP_START}", index);
        if(pos_start == std::string::npos)
        {
            std::cerr << "can not find " << index << " th loop start mark, over range";
            return -2;
        }
        loop_start = findAround(document_xml, "<w:p>", pos_start, UP);
        if(loop_start == std::string::npos || loop_start == -1) return -3;
        int pos_end = myFind(document_xml, "${LP_END}", index);
        if(pos_end == std::string::npos)
        {
            std::cerr << "can not find " << index << " th loop end mark, over range" << std::endl;
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
                readXml(rep_model[i], ":/new/xml_models/xml_models/image.xml");
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
                    int rid = addImage(("insert_image" + std::to_string((++ image_sn))), rep_info[i][info_index ++]);
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

                        replaceText(tmp_mark, rep_info[i][info_index ++], tmp);
                    }
                }
                rep_res.append(tmp);
            }
        }

        document_xml.replace(loop_start, loop_end - loop_start, rep_res);
    }
    
    writeXml(document_xml, (cache_path + "/word/document.xml"));
    return 0;
}

int WordOp::addInfoRecursive(std::vector<int> indexs, std::vector<std::vector<std::vector<std::string>>> rep_infos)
{
    QDir tempdir(QString::fromStdString(cache_path + "/word/media"));
    int image_sn = tempdir.count() - 2;    
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
            std::cerr << "Error rep_info empty" << std::endl;
            return -1;
        }
        int pos_start = myFind(document_xml, "${LP_START}", index);
        if(pos_start == std::string::npos)
        {
            std::cerr << "can not find " << index << " th loop start mark, over range";
            return -2;
        }
        loop_start = findAround(document_xml, "<w:p>", pos_start, UP);
        if(loop_start == std::string::npos || loop_start == -1) return -3;
        int pos_end = myFind(document_xml, "${LP_END}", index);
        if(pos_end == std::string::npos)
        {
            std::cerr << "can not find " << index << " th loop end mark, over range" << std::endl;
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
                readXml(rep_model[i], ":/new/xml_models/xml_models/image.xml");
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
                    int rid = addImage(("insert_image" + std::to_string((++ image_sn))), rep_info[i][info_index ++]);
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

                        replaceText(tmp_mark, rep_info[i][info_index ++], tmp);
                    }
                }
                rep_res.append(tmp);
            }
        }

        document_xml.replace(loop_start, loop_end - loop_start, rep_res);
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
            std::cerr << "The " << index << "th TB_START not found" << std::endl;
            return -1;
        }
        loop_start = findAround(document_xml, "<w:tr>", pos_start, UP);
        if(loop_start == std::string::npos || loop_start == -1)
        {
            std::cerr << "The " << index << "th loop_start not found" << std::endl;
            return -2;
        }

        int pos_end = myFind(document_xml, "${TB_END}", index);
        if(pos_end == std::string::npos)
        {
            std::cerr << "The " << index << "th TB_END not found" << std::endl;
            return -1;
        }
        loop_end = findAround(document_xml, "</w:tr>", pos_end, DOWN);
        if(loop_end == std::string::npos || loop_end == -1)
        {
            std::cerr << "The " << index << "th loop_end not found" << std::endl;
            return -2;
        }
        loop_end += 7 + pos_end;
        tmp = document_xml.substr(pos_start, pos_end - pos_start);
        pos_start = myFind(tmp, "<w:tr>", 1);
        pos_end = myFind(tmp, "<w:tr>", -1);

        analyzeXml(rep_model, tmp.substr(pos_start, pos_end - pos_start), "w:tr");

        for(int i = 0; i < rep_info.size(); ++ i)
        {
            std::vector<std::string> tmp_rep_model = rep_model;
            for(int j = 0; j < rep_info[i].size(); ++ j)
            {
                std::string rep_mark = "${TB_ARG" + std::to_string(j + 1) + "}";
                for(int k = 0; k < tmp_rep_model.size(); ++ k)
                {
                    if(rep_model[k].find(rep_mark) != std::string::npos)
                    {
                        replaceText(rep_mark, rep_info[i][j], tmp_rep_model[k]);
                    }
                }
            }
            for(int j = 0; j < rep_model.size(); ++ j) rep_res.append(tmp_rep_model[j]);
        }
        document_xml.replace(loop_start, loop_end - loop_start, rep_res);
    }

    return 0;
}


void WordOp::readXml(std::string &xml_file, std::string filepath)
{
    xml_file.clear();
    QFile file(QString::fromStdString(filepath));
    QString res;
    file.open(QIODevice::ReadOnly);
    if(!file.isOpen())
    {
        std::cerr << "Error: Could not open file: " << filepath << std::endl;
        return;
    }
    while(!file.atEnd())
    {
        QString tmp_res = file.readLine();
        res.append(tmp_res + "\n");
    }
    xml_file = res.toStdString();

    file.close();
}

void WordOp::readList(std::vector<std::vector<std::string>> &str_list, std::string filepath)
{

    QFile file(QString::fromStdString(filepath));
    file.open(QIODevice::ReadOnly);
    if(!file.isOpen())
    {
        std::cerr << "Error: Could not open file: " << filepath << std::endl;
        return;
    }
    std::vector<std::string> tmp_vec;
    while(!file.atEnd())
    {
        QString tmp_line = file.readLine();
        if(tmp_line != "\n")
        {
            tmp_vec.push_back(tmp_line.toStdString());
            tmp_vec.back().pop_back();
        }
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

    for(auto a : tmp_vec)
    {
        if(!a.empty())
        {
            str_list.push_back(tmp_vec);
            break;
        }
    }
    tmp_vec.clear();
    file.close();
    return;
}

int WordOp::myFind(std::string src, std::string des, int index)
{
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
        std::cerr << "findAround argument error" << std::endl;
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
}

int WordOp::addImage(std::string mark, std::string replaced_image)
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
        std::cerr << "Error can not open image " << replaced_image << std::endl;
        return -1;
    }
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