#include "wordop.h"

QString WordOp::line_feed_head = "$</w:t$>$</w:r$>$</w:p$>$<w:p$>$<w:r$>";
QString WordOp::line_feed_tail = "$<w:t$>";

/*---------------
Public Functions
---------------*/

WordOp::WordOp(QString filepath)
{
    origin_work_path = QDir::currentPath();
    QDir::setCurrent(QCoreApplication::applicationDirPath());
    cache_path = "";
    des_path = "";
    this->filepath = filepath;
    image_rels = "<Relationship Id=\"${rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"${target}\"/>";
    qDebug() << "A WordOp instance has been created.";
}

WordOp::WordOp(QString filepath, QString des_path)
{
    origin_work_path = QDir::currentPath();
    QDir::setCurrent(QCoreApplication::applicationDirPath());
    cache_path = "";
    this->filepath = filepath;
    this->des_path = des_path;
    image_rels = "<Relationship Id=\"${rId}\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"${target}\"/>";
    qDebug() << "A WordOp instance has been created.";
}

WordOp::~WordOp()
{
    QDir::setCurrent(origin_work_path);
    qDebug() << "A WordOp instance has been destroyed.";
}

void WordOp::open(QString filepath)
{
    if(!(filepath.isEmpty())) this->filepath = filepath;
    if(!des_path.isEmpty()) cache_path = FileOp::unzipFolder(this->filepath, this->des_path);
    else cache_path = FileOp::unzipFolder(this->filepath);

    qDebug() << "filepath: " << filepath << " des_path: " << des_path;
    readXml(document_xml, (cache_path + "/word/document.xml"));
}

void WordOp::open(QString filepath, QString des_path)
{
    qDebug() << "filepath: " << filepath;
    if(!filepath.isEmpty()) this->filepath = filepath;
    cache_path = FileOp::unzipFolder(this->filepath, des_path);
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

int WordOp::replaceText(QString mark, QString replaced_text, QString &rep_str)
{

    mark.replace("<", "&lt;");
    mark.replace(">", "&gt;");
    for(int i = 0; i < mark.size() - 3; ++ i)
    {
        if(mark[i] == '&' && mark.mid(i, 4) != "&lt;" && mark.mid(i, 4) != "&gt;")
        mark.replace(i, 1, "&amp;");
    }
        
    replaced_text.replace("<", "&lt;");
    replaced_text.replace(">", "&gt;");
    replaced_text.replace("$&lt;", "<");
    replaced_text.replace("$&gt;", ">");

    for(int i = 0; i < replaced_text.size() - 3; ++ i)
    {
        if(replaced_text[i] == '&' && replaced_text.mid(i, 4) != "&lt;" && replaced_text.mid(i, 4) != "&gt;")
        replaced_text.replace(i, 1, "&amp;");
    }

    if(rep_str == "")
    {
        
        
        if(document_xml.isEmpty()) return -1;
        int p1 = myFind(document_xml, "<w:p>", 1);
        int p2 = myFind(document_xml, "</w:p>", -1);

        QString tmp = document_xml.mid(p1, p2 + 6 - p1);

        std::vector<QString> ana;

        analyzeXml(ana, tmp, "w:p");

        int pos = -1;
        while((pos = document_xml.indexOf(mark, pos + 1)) != -1)
        {
            document_xml.replace(pos, mark.length(), replaced_text);
        }
    }
    else
    {
        if(rep_str.isEmpty()) return -1;
        int p1 = myFind(rep_str, "<w:p>", 1);
        int p2 = myFind(rep_str, "</w:p>", -1);

        QString tmp = rep_str.mid(p1, p2 + 6 - p1);

        rep_str.replace(mark, replaced_text);
    }
    return 0;
}

int WordOp::replaceText(Info info)
{
    std::map<QString, QString> mark_to_replacements_map = info.getLabelText();
    std::vector<QString> marks, replaced_texts;
    for(auto a = mark_to_replacements_map.begin(); a != mark_to_replacements_map.end(); ++a)
    {
        marks.push_back((*a).first);
        replaced_texts.push_back((*a).second);
    }
    return replaceText(marks, replaced_texts);
}

int WordOp::replaceText(std::vector<QString> marks, std::vector<QString> replaced_texts)
{
    QString rep_str = "";
    return replaceText(marks, replaced_texts, rep_str);
}

int WordOp::replaceText(std::vector<QString> marks, std::vector<QString> replaced_texts, QString &rep_str)
{
    if(marks.size() != replaced_texts.size())
    {
        qDebug() << "Mark size does not match ReplaceText size";
        return -1;
    }
    for(int i = 0; i < marks.size(); ++ i)
    {
        replaceText(marks[i], replaced_texts[i], rep_str);
    }
    return 0;
}

int WordOp::replaceText(QString mark_file, QString replaced_text_file)
{
    std::vector<std::vector<QString>> rep_info, mark_info;
    std::vector<QString> marks, reps;
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
        qDebug() << "Items in mark file does not equal to items in replaced_text file";
        return -1;
    }
    replaceText(marks, reps);
    return 0;
}

int WordOp::replaceImage(std::vector<QString> marks, std::vector<QString> replace_image_paths)
{
    if(marks.size() != replace_image_paths.size())
    {
        qDebug() << "count of marks does not equal to replace_image_paths";
        return -1;
    }
    for(int i = 0; i < marks.size(); ++ i)
    {
        QString mark = marks[i];
        QString replace_image_path = replace_image_paths[i];
        int mark_pos = myFind(document_xml, mark, 1);
        QDir dir((cache_path + "/word/media"));
        cv::Mat tmp_image = cv::imread(replace_image_path.toStdString());
        int image_height = tmp_image.size().height * NORMAL_IMAGE_SIZE_TIMES;
        int image_width = tmp_image.size().width * NORMAL_IMAGE_SIZE_TIMES;
        int rid = addImage(replace_image_path);
        QString img_model;
        readXml(img_model, IMAGE_MODEL_PATH);
        replaceText("${ID}", QString::number(rid), img_model);
        replaceText("${NAME}", ("image" + QString::number(rid)), img_model);
        replaceText("${IMAGE_SN}", ("rId" + QString::number(rid)), img_model);
        replaceText("${H}", "align", img_model);
        replaceText("${HARG}", "center", img_model);
        replaceText("${V}", "posOffset", img_model);
        replaceText("${VARG}", "653", img_model);
        replaceText("${CX}", QString::number(image_width), img_model);
        replaceText("${CY}", QString::number(image_height), img_model);

        while(mark_pos != -1)
        {
            int rep_start = findAround(document_xml, "<w:p>", mark_pos, UP);
            if(rep_start == -1 || rep_start == -1)
            {
                qDebug() << "Error fail to indexOf <w:p> around pos " << mark_pos << " using UP";
                return -2;
            }
            int rep_end = findAround(document_xml, "</w:p>", mark_pos, DOWN);
            if(rep_end == -1 || rep_end == -1)
            {
                qDebug() << "Error fail to indexOf </w:p> around pos " << mark_pos << " using DOWN";
                return -2;
            }
            rep_end += mark_pos + 6;
            document_xml.replace(rep_start, rep_end - rep_start, img_model);

            mark_pos = myFind(document_xml, mark, 1);
        }
    }
    return 0;
}

int WordOp::replaceImageFromMat(std::vector<QString> marks, std::vector<cv::Mat> replace_mat_images)
{
    if(marks.size() != replace_mat_images.size())
    {
        qDebug() << "count of marks does not equal to mats";
        return -1;
    }
    for(int i = 0; i < marks.size(); ++ i)
    {
        QString mark = marks[i];
        int mark_pos = myFind(document_xml, mark, 1);
        QDir dir((cache_path + "/word/media"));
        int rid = addImageFromMat(replace_mat_images[i]);
        if(rid == -1) continue;
        int image_height = replace_mat_images[i].size().height * NORMAL_IMAGE_SIZE_TIMES;
        int image_width = replace_mat_images[i].size().width * NORMAL_IMAGE_SIZE_TIMES;
        QString img_model;
        readXml(img_model, IMAGE_MODEL_PATH);
        replaceText("${ID}", QString::number(rid), img_model);
        replaceText("${NAME}", ("image" + QString::number(rid)), img_model);
        replaceText("${IMAGE_SN}", ("rId" + QString::number(rid)), img_model);
        replaceText("${H}", "align", img_model);
        replaceText("${HARG}", "center", img_model);
        replaceText("${V}", "posOffset", img_model);
        replaceText("${VARG}", "653", img_model);
        replaceText("${CX}", QString::number(image_width), img_model);
        replaceText("${CY}", QString::number(image_height), img_model);
        while(mark_pos != -1)
        {
            int rep_start = findAround(document_xml, "<w:p>", mark_pos, UP);
            if(rep_start == -1 || rep_start == -1)
            {
                qDebug() << "Error fail to indexOf <w:p> around pos " << mark_pos << " using UP";
                return -2;
            }
            int rep_end = findAround(document_xml, "</w:p>", mark_pos, DOWN);
            if(rep_end == -1 || rep_end == -1)
            {
                qDebug() << "Error fail to indexOf </w:p> around pos " << mark_pos << " using DOWN";
                return -2;
            }
            rep_end += mark_pos + 6;
            document_xml.replace(rep_start, rep_end - rep_start, img_model);

            mark_pos = myFind(document_xml, mark, 1);
        }
    }
    
    return 0;
}

int WordOp::replaceImageFromFile(QString mark_file_path, QString replace_image_file_path)
{
    std::vector<std::vector<QString>> tmp_marks, tmp_replace_image_paths;
    std::vector<QString> marks, replace_image_paths;
    readList(tmp_marks, mark_file_path);
    readList(tmp_replace_image_paths, replace_image_file_path);
    for(int i = 0; i < tmp_marks.size(); ++ i)
    {
        for(int j = 0; j < tmp_marks[i].size(); ++ j)
        {
            marks.push_back(tmp_marks[i][j]);
        }
    }
    for(int i = 0; i < tmp_replace_image_paths.size(); ++ i)
    {
        for(int j = 0; j < tmp_replace_image_paths[i].size(); ++ j)
        {
            replace_image_paths.push_back(tmp_replace_image_paths[i][j]);
        }
    }

    if(replace_image_paths.size() != marks.size())
    {
        qDebug() << "Error from replaceImageFromFile: rep_images size does not match marks";
        return -1;
    }

    replaceImage(marks, replace_image_paths);

    return 0;
}

int WordOp::addInfoRecursive(std::vector<int> indexs, std::vector<Info> &infos)
{
    QString tmp;
    for(int dex = indexs.size() - 1; dex >= 0; dex --)
    {
        int loop_start, loop_end, tmp_pos, index;
        std::vector<QString> rep_model;
        std::vector<std::vector<QString>> rep_info;
        QString rep_res;

        loop_start = loop_end = tmp_pos = index = 0;
        rep_model.clear();
        rep_info.clear();
        rep_res.clear();
        index = indexs[dex];

        int pos_start = myFind(document_xml, LP_START_MARK, index);
        if(pos_start == -1)
        {
            qDebug() << "can not indexOf " << index << " th loop start mark, over range";
            return -2;
        }
        loop_start = findAround(document_xml, "<w:p>", pos_start, UP);
        if(loop_start == -1 || loop_start == -1) return -3;
        int pos_end = myFind(document_xml, LP_END_MARK, index);
        if(pos_end == -1)
        {
            qDebug() << "can not indexOf " << index << " th loop end mark, over range";
            return -3;
        }

        loop_end = pos_end + findAround(document_xml, "</w:p>", pos_end, DOWN);
        if(loop_end == -1 || loop_end == -1) return -3;
        loop_end += 6;
        tmp = document_xml.mid(pos_start, pos_end - pos_start);
        pos_start = myFind(tmp, "<w:p>", 1);
        pos_end = myFind(tmp, "<w:p>", -1);

        analyzeXml(rep_model, tmp.mid(pos_start, pos_end - pos_start), "w:p");

        for(int i = 0; i < infos.size(); ++ i)
        {
            std::map<QString, QString> label_to_text_map = infos[i].getLabelText();
            std::map<QString, cv::Mat> label_to_mat_map = infos[i].getLabelImage();
            for(int j = 0; j < rep_model.size(); ++ j)
            {
                st : 
                if(j < rep_model.size()) tmp = rep_model[j];
                else break;

                for(auto a : label_to_mat_map)
                {
                    if(tmp.indexOf(a.first) != -1)
                    {
                        tmp.clear();
                        readXml(tmp, IMAGE_MODEL_PATH);
                        int rid = addImageFromMat(a.second);
                        if(rid == -1)
                        {
                            ++ j;
                            goto st;
                        }
                        int image_height = a.second.size().height * NORMAL_IMAGE_SIZE_TIMES;
                        int image_width = a.second.size().width * NORMAL_IMAGE_SIZE_TIMES;
                        replaceText("${ID}", QString::number(rid), tmp);
                        replaceText("${NAME}", ("image" + QString::number(rid)), tmp);
                        replaceText("${IMAGE_SN}", ("rId" + QString::number(rid)), tmp);
                        replaceText("${H}", "align", tmp);
                        replaceText("${HARG}", "center", tmp);
                        replaceText("${V}", "posOffset", tmp);
                        replaceText("${VARG}", "653", tmp);
                        replaceText("${CX}", QString::number(image_width), tmp);
                        replaceText("${CY}", QString::number(image_height), tmp);
                        rep_res.append(tmp);
                        ++ j;
                        goto st;
                    }
                }
                for(auto a : label_to_text_map)
                {
                    if(tmp.indexOf(a.first) != -1)
                    {
                        replaceText(a.first, a.second, tmp);
                        rep_res.append(tmp);
                        ++ j;

                        goto st;
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

int WordOp::addTableRows(std::vector<int> indexs, std::vector<Info> &infos)
{
    QString tmp;
    for(int dex = indexs.size() - 1; dex >= 0; -- dex)
    {
        int loop_start, loop_end, tmp_pos, index;
        std::vector<QString> rep_model;
        std::vector<std::vector<QString>> rep_info;
        QString rep_res;

        loop_start = loop_end = tmp_pos = index = 0;
        rep_model.clear();
        rep_res.clear();
        index = indexs[dex];
        int pos_start = myFind(document_xml, TABLE_LOOP_START_MARK, index);
        if(pos_start == -1)
        {
            qDebug() << "The " << index << "th TB_START not found";
            return -1;
        }
        loop_start = findAround(document_xml, "<w:tr>", pos_start, UP);
        if(loop_start == -1 || loop_start == -1)
        {
            qDebug() << "The " << index << "th loop_start not found";
            return -2;
        }

        int pos_end = myFind(document_xml, TABLE_LOOP_END_MARK, index);
        if(pos_end == -1)
        {
            qDebug() << "The " << index << "th TB_END not found";
            return -1;
        }
        loop_end = findAround(document_xml, "</w:tr>", pos_end, DOWN);
        if(loop_end == -1 || loop_end == -1)
        {
            qDebug() << "The " << index << "th loop_end not found";
            return -2;
        }
        loop_end += 7 + pos_end;
        tmp = document_xml.mid(pos_start, pos_end - pos_start);
        pos_start = myFind(tmp, "<w:tr>", 1);
        pos_end = myFind(tmp, "<w:tr>", -1);

        analyzeXml(rep_model, tmp.mid(pos_start, pos_end - pos_start), "w:tr");
        for(int i = 0; i < infos.size(); ++ i)
        {
            std::map<QString, QString> label_to_text_map = infos[i].getLabelText();
            std::map<QString, cv::Mat> label_to_mat_map = infos[i].getLabelImage();
            for(int j = 0; j < rep_model.size(); ++ j)
            {
                tmp = rep_model[j];
                for(auto a : label_to_mat_map)
                {
                    tmp_pos = tmp.indexOf(a.first);
                    if(tmp_pos != -1)
                    {                        
                        int p1 = findAround(tmp, "<w:p>", tmp_pos, UP);
                        int p2 = tmp_pos + findAround(tmp, "</w:p>", tmp_pos, DOWN);
                        QString image_model;
                        readXml(image_model, IMAGE_MODEL_PATH);
                        tmp.replace(p1, p2 - p1 + 6, image_model);
                        int rid = addImageFromMat(a.second);
                        if(rid == -1) continue;
                        int image_height = a.second.size().height * TABLE_IMAGE_SIZE_TIMES;
                        int image_width = a.second.size().width * TABLE_IMAGE_SIZE_TIMES;

                        replaceText("${ID}", QString::number(rid), tmp);
                        replaceText("${NAME}", ("image" + QString::number(rid)), tmp);
                        replaceText("${IMAGE_SN}", ("rId" + QString::number(rid)), tmp);
                        replaceText("${H}", "align", tmp);
                        replaceText("${HARG}", "center", tmp);
                        replaceText("${V}", "posOffset", tmp);
                        replaceText("${VARG}", "653", tmp);
                        replaceText("${CX}", QString::number(image_width), tmp);
                        replaceText("${CY}", QString::number(image_height), tmp);
                    }
                }
                for(auto a : label_to_text_map)
                {
                    if(tmp.indexOf(a.first) != -1)
                    {
                        replaceText(a.first, a.second, tmp);
                    }
                }
                rep_res.append(tmp);
            }
            
        }
        document_xml.replace(loop_start, loop_end - loop_start, rep_res);
    }

    return 0;
}

/*---------------
Private Functions
---------------*/

void WordOp::readXml(QString &xml_file, QString filepath)
{
    xml_file.clear();
    QFile file((filepath));
    QString res;
    file.open(QIODevice::ReadOnly);
    if(!file.isOpen())
    {
        qDebug() << "Error: Could not open file: " << filepath;
        return;
    }
    while(!file.atEnd())
    {
        QString tmp_res = file.readLine();
        res.append(tmp_res);
    }
    xml_file = res;

    file.close();
}

void WordOp::readList(std::vector<std::vector<QString>> &str_list, QString filepath)
{

    QFile file((filepath));
    file.open(QIODevice::ReadOnly);
    if(!file.isOpen())
    {
        qDebug() << "Error: Could not open file: " << filepath;
        return;
    }
    std::vector<QString> tmp_vec;
    while(!file.atEnd())
    {
        QString tmp_line = file.readLine();
        if(tmp_line != "\n")
        {
            tmp_line.remove(tmp_line.length() - 1, 1);
            tmp_vec.push_back(tmp_line);
        }
        else
        {
            for(auto a : tmp_vec)
            {
                if(!a.isEmpty())
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
        if(!a.isEmpty())
        {
            str_list.push_back(tmp_vec);
            break;
        }
    }
    tmp_vec.clear();
    file.close();
    return;
}

int WordOp::myFind(QString src, QString des, int index)
{
    int _index = 1;
    int pos = 0;
    pos = src.indexOf(des);
    if(index == -1)
    {
        pos = src.lastIndexOf(des);
    }
    else
    {
        while(_index != index)
        {
            pos = src.indexOf(des, pos + des.size());
            ++ _index;
            if(pos == -1) break;
        }
    }
    
    return pos;
}

int WordOp::findAround(QString src, QString des, int cur_pos, int direction)
{
    QString tmp;
    int ret;
    if(direction == UP)
    {
        tmp = src.mid(0, cur_pos);
        ret = myFind(tmp, des, -1);
    }
    else if(direction == DOWN)
    {
        tmp = src.mid(cur_pos);
        ret = myFind(tmp, des, 1);
    }
    else
    {
        ret = -1;
        qDebug() << "findAround argument error";
    }

    return ret;
}

void WordOp::analyzeXml(std::vector<QString> &analysis_xml, QString origin_xml, QString flag)
{
    analysis_xml.clear();
    int pos_start, pos_end, index = 1;
    int tmp_pos_before = 0;
    int tmp_pos_after = 0;
    pos_start = myFind(origin_xml, ("<" + flag + ">"), index);
    pos_end = myFind(origin_xml, ("</" + flag + ">"), index);
    while(pos_start != -1)
    {
        analysis_xml.push_back(origin_xml.mid(pos_start, pos_end + ("</" + flag + ">").size() - pos_start));
        tmp_pos_before = analysis_xml.back().indexOf("<w:t>", 0);
        while(tmp_pos_before != -1)
        {
            tmp_pos_after = analysis_xml.back().indexOf("</w:t>", tmp_pos_before + 6);

            int pos_front = analysis_xml.back().indexOf('<', tmp_pos_before + 6);
            int pos_back = analysis_xml.back().indexOf('>', tmp_pos_before + 6);

            if(pos_back < tmp_pos_after)
            {
                analysis_xml.back().replace(pos_front, 1, "&lt;");
                analysis_xml.back().replace(pos_back, 1, "&gt;");
            }
            
            tmp_pos_before = analysis_xml.back().indexOf("<w:t>", tmp_pos_before + 6);
        }
        for(int i = 0; i < analysis_xml.back().size() - 3; ++ i)
        {
            if(analysis_xml.back()[i] == '&' && analysis_xml.back().mid(i, 4) != "&lt;" && analysis_xml.back().mid(i, 4) != "&gt;")
            analysis_xml.back().replace(i, 1, "&amp;");
        }
        
        ++index;
        pos_start = myFind(origin_xml, ("<" + flag + ">"), index);
        pos_end = myFind(origin_xml, ("</" + flag + ">"), index);
    }
}

void WordOp::writeXml(QString &xml_file, QString filepath)
{
	QFile file((filepath));
	file.open(QIODevice::WriteOnly);
	file.write(xml_file.toUtf8());
	file.close();
}

int WordOp::addImage(QString replace_image_path)
{
    QDir tempdir((cache_path + "/word/media"));
    int image_sn = tempdir.count() - 2;    
    QString mark = ("image" + QString::number((image_sn + 1)));
    QFile f_mark, f_image;
    int rId_sn = 1;
	QDir tmp_dir((cache_path + "/word/media/"));
	if(!tmp_dir.exists()) tmp_dir.mkpath((cache_path + "/word/media/"));

    mark += ".png";
    readXml(doc_rels_xml, (cache_path + "/word/_rels/document.xml.rels"));
    int pos;
    while(true)
    {
        pos = doc_rels_xml.indexOf(("rId" + QString::number(rId_sn)));
        if(pos == -1)
        {
            QString new_image_rels = image_rels;
            pos = new_image_rels.indexOf("${rId}");
            new_image_rels.replace(pos, 6, ("rId" + QString::number(rId_sn)));
            pos = new_image_rels.indexOf("${target}");
            new_image_rels.replace(pos, 9, ("media/" + mark));

            pos = doc_rels_xml.indexOf("</Relationships>");
            doc_rels_xml.replace(pos, 16, new_image_rels + "</Relationships>\n");

            writeXml(doc_rels_xml, (cache_path + "/word/_rels/document.xml.rels"));
            break;
        }
        ++ rId_sn;
    }

	if (!QFile::exists((replace_image_path)))
	{
		qDebug() << "Error can not open image " << replace_image_path;
		return -1;
	}

	f_mark.setFileName((cache_path + "/word/media/" + mark));
	f_image.setFileName((replace_image_path));
    f_mark.open(QIODevice::ReadWrite);
	f_image.open(QIODevice::ReadWrite);
    
	QByteArray b_array = f_image.readAll();
	f_mark.write(b_array);

    f_mark.close();
    f_image.close();
    return rId_sn;
}

int WordOp::addImageFromMat(cv::Mat replace_mat_image)
{
    if(replace_mat_image.empty())
	{
		qDebug() << "Error replace_mat_image is isEmpty ";
		return -1;
	}

    QDir tempdir((cache_path + "/word/media"));
    int image_sn = tempdir.count() - 2;    
    QString mark = ("image" + QString::number((image_sn + 1)));
    int rId_sn = 1;
	QDir tmp_dir((cache_path + "/word/media/"));
	if(!tmp_dir.exists()) tmp_dir.mkpath((cache_path + "/word/media/"));

    mark += ".png";
    readXml(doc_rels_xml, (cache_path + "/word/_rels/document.xml.rels"));
    int pos;
    while(true)
    {
        pos = doc_rels_xml.indexOf(("rId" + QString::number(rId_sn)));
        if(pos == -1)
        {
            QString new_image_rels = image_rels;
            pos = new_image_rels.indexOf("${rId}");
            new_image_rels.replace(pos, 6, ("rId" + QString::number(rId_sn)));
            pos = new_image_rels.indexOf("${target}");
            new_image_rels.replace(pos, 9, ("media/" + mark));

            pos = doc_rels_xml.indexOf("</Relationships>");
            doc_rels_xml.replace(pos, 16, new_image_rels + "</Relationships>\n");

            writeXml(doc_rels_xml, (cache_path + "/word/_rels/document.xml.rels"));
            break;
        }
        ++ rId_sn;
    }

    bool ret = cv::imwrite((cache_path + "/word/media/" + mark).toStdString(), replace_mat_image);
    if(ret == false)
    {
        qDebug() << "Fail to write replace_mat_image to " << (cache_path + "/word/media/" + mark);
        return -2;
    }
    else return rId_sn;
}