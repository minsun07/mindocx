#include <Aspose.Words.Cpp/Document.h>                                                         // word ������ ���� �� ó���� ���� Ŭ����
#include <Aspose.Words.Cpp/DocumentBuilder.h>                                                  // ������ �ý�Ʈ�� ��Ҹ� �߰��ϴ� Ŭ����
#include <Aspose.Words.Cpp/Drawing/Shape.h>                                                    // �׸�, ������ ���� �� ��ġ�� ���� Ŭ����
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/WrapType.h>                                                 
#include <Aspose.Words.Cpp/Font.h>                                                              // ��Ʈ �������� ���� Ŭ����
#include <iostream>
#include <string>
#include <curl/curl.h>
#include <json/json.h>

using namespace Aspose::Words;                                                                 // ���� ���� �� ������ ���� Ŭ������ ����
using namespace Aspose::Words::Drawing;                                                        // ������ ���� �� �׷��� ��Ҹ� �ٷ�
using namespace System;                                                                        // Aspose���̺귯������ ���Ǵ� �ֿ� ������ Ÿ��(SharedPtr ��)�� ����
using namespace std;

// cURL ���� �����͸� ó���ϴ� �ݹ� �Լ�
size_t WriteCallback(void* contents, size_t size, size_t nmemb, void* userp) {
    ((string*)userp)->append((char*)contents, size * nmemb);
    return size * nmemb;
} 

int main() {
    try {
        // input.docx ���� �б�
        System::SharedPtr<Document> doc = System::MakeObject<Document>(u"C:/2024-CPP-project/mindocx/x64/Debug/input.docx");
        System::String extractedText = doc->GetText(); // �������� �ؽ�Ʈ ����
        wcout << L"����� �ؽ�Ʈ: " << extractedText << endl;

        // GPT API�� �ؽ�Ʈ ��� ��û
        string api_key = "sk-your-openai-api-key"; // ���⿡ API Ű �Է�
        string api_url = "https://api.openai.com/v1/chat/completions";

        // ��û�� JSON ������ �غ�
        Json::Value root;
        root["model"] = "gpt-3.5-turbo";
        root["messages"] = Json::arrayValue;
        Json::Value message;
        message["role"] = "user";
        message["content"] = "Summarize the following text: " + extractedText.ToUtf8String();
        root["messages"].append(message);

        // JSON �����͸� ���ڿ��� ��ȯ
        Json::StreamWriterBuilder writer;
        string request_data = Json::writeString(writer, root);

        // cURL�� ����Ͽ� HTTP ��û ������
        CURL* curl;
        CURLcode res;
        string response_data;

        curl_global_init(CURL_GLOBAL_DEFAULT);
        curl = curl_easy_init();
        if (curl) {
            curl_easy_setopt(curl, CURLOPT_URL, api_url.c_str());
            curl_easy_setopt(curl, CURLOPT_POSTFIELDS, request_data.c_str());

            // ��� ���� (Authorization)
            struct curl_slist* headers = NULL;
            headers = curl_slist_append(headers, "Content-Type: application/json");
            headers = curl_slist_append(headers, ("Authorization: Bearer " + api_key).c_str());
            curl_easy_setopt(curl, CURLOPT_HTTPHEADER, headers);

            // ���� �����͸� ó���ϴ� �ݹ� ����
            curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, WriteCallback);
            curl_easy_setopt(curl, CURLOPT_WRITEDATA, &response_data);

            // ��û ������
            res = curl_easy_perform(curl);

            if (res != CURLE_OK) {
                cerr << "cURL failed: " << curl_easy_strerror(res) << endl;
            }
            else {
                cout << "API ����: " << response_data << endl;
            }

            curl_slist_free_all(headers);
            curl_easy_cleanup(curl);
        }

        curl_global_cleanup();

        // ���信�� ���� �ؽ�Ʈ ����
        Json::CharReaderBuilder reader;
        Json::Value response_json;
        string errs;
        string summarized_text;

        // JSON �����͸� ��Ʈ���� �����ϰ�, ��Ʈ���� ����
        istringstream response_stream(response_data);

        if (Json::parseFromStream(reader, response_stream, &response_json, &errs)) {
            summarized_text = response_json["choices"][0]["message"]["content"].asString();
        }
        else {
            cerr << "JSON ���� �Ľ� ����: " << errs << endl;
            return 1;
        }

        // ���� �ؽ�Ʈ�� ���� ������ �߰�
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"\n--- ��� ---");
        builder->Writeln(u"����: " + System::String(summarized_text.c_str()));

        // ���� �ؽ�Ʈ�� ���Ե� ������ ���� input.docx�� �����
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/input.docx");

        cout << "���� ������ input.docx�� ���� �Ϸ�!" << endl;

    }
    catch (const System::Exception& e) {
        cerr << "Aspose ���� �߻�: " << e->get_Message().ToUtf8String() << endl;
    }
    catch (const std::exception& e) {
        cerr << "std::exception �߻�: " << e.what() << endl;
    }
    catch (...) {
        cerr << "�� �� ���� ���� �߻�" << endl;
    }
}