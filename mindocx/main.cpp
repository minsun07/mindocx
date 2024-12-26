#include <Aspose.Words.Cpp/Document.h>                                                         // word 문서의 생성 및 처리를 위한 클래스
#include <Aspose.Words.Cpp/DocumentBuilder.h>                                                  // 문서에 택스트나 요소를 추가하는 클래스
#include <Aspose.Words.Cpp/Drawing/Shape.h>                                                    // 그림, 도형을 삽입 및 위치를 위한 클래스
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/WrapType.h>                                                 
#include <Aspose.Words.Cpp/Font.h>                                                              // 폰트 디자인을 위한 클래스
#include <iostream>
#include <string>
#include <curl/curl.h>
#include <json/json.h>

using namespace Aspose::Words;                                                                 // 문서 생성 및 조작을 위한 클래스가 포함
using namespace Aspose::Words::Drawing;                                                        // 문서에 도형 및 그래픽 요소를 다룸
using namespace System;                                                                        // Aspose 라이브러리에서 사용되는 주요 데이터 타입(SharedPtr 등)을 포함
using namespace std;

// cURL 응답 데이터를 처리하는 콜백 함수
size_t WriteCallback(void* contents, size_t size, size_t nmemb, void* userp) {
    ((string*)userp)->append((char*)contents, size * nmemb);
    return size * nmemb;
}

int main() {
    try {
        // input.docx 문서 읽기
        System::SharedPtr<Document> doc = System::MakeObject<Document>(u"C:/2024-CPP-project/mindocx/x64/Debug/input.docx");
        System::String extractedText = doc->GetText(); // 문서에서 텍스트 추출
        wcout << L"추출된 텍스트: " << extractedText << endl;

        // GPT API에 텍스트 요약 요청
        string api_key = "sk-your-openai-api-key"; // 여기에 API 키 입력
        string api_url = "https://api.openai.com/v1/chat/completions";

        // 요청할 JSON 데이터 준비
        Json::Value root;
        root["model"] = "gpt-3.5-turbo";
        root["messages"] = Json::arrayValue;
        Json::Value message;
        message["role"] = "user";
        message["content"] = "Summarize the following text: " + extractedText.ToUtf8String();
        root["messages"].append(message);

        // JSON 데이터를 문자열로 변환
        Json::StreamWriterBuilder writer;
        string request_data = Json::writeString(writer, root);

        // cURL을 사용하여 HTTP 요청 보내기
        CURL* curl;
        CURLcode res;
        string response_data;

        curl_global_init(CURL_GLOBAL_DEFAULT);
        curl = curl_easy_init();
        if (curl) {
            curl_easy_setopt(curl, CURLOPT_URL, api_url.c_str());
            curl_easy_setopt(curl, CURLOPT_POSTFIELDS, request_data.c_str());

            // 헤더 설정 (Authorization)
            struct curl_slist* headers = NULL;
            headers = curl_slist_append(headers, "Content-Type: application/json");
            headers = curl_slist_append(headers, ("Authorization: Bearer " + api_key).c_str());
            curl_easy_setopt(curl, CURLOPT_HTTPHEADER, headers);

            // 응답 데이터를 처리하는 콜백 설정
            curl_easy_setopt(curl, CURLOPT_WRITEFUNCTION, WriteCallback);
            curl_easy_setopt(curl, CURLOPT_WRITEDATA, &response_data);

            // 요청 보내기
            res = curl_easy_perform(curl);

            if (res != CURLE_OK) {
                cerr << "cURL failed: " << curl_easy_strerror(res) << endl;
            }
            else {
                cout << "API 응답: " << response_data << endl;
            }

            curl_slist_free_all(headers);
            curl_easy_cleanup(curl);
        }

        curl_global_cleanup();

        // 응답에서 요약된 텍스트 추출
        Json::CharReaderBuilder reader;
        Json::Value response_json;
        string errs;
        string summarized_text;

        // JSON 데이터를 스트림에 저장하고, 스트림을 전달
        istringstream response_stream(response_data);

        if (Json::parseFromStream(reader, response_stream, &response_json, &errs)) {
            summarized_text = response_json["choices"][0]["message"]["content"].asString();
        }
        else {
            cerr << "JSON 응답 파싱 실패: " << errs << endl;
            return 1;
        }

        // 요약된 텍스트를 기존 문서에 추가
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);
        builder->Writeln(u"\n--- 요약 ---");
        builder->Writeln(u"내용: " + System::String(summarized_text.c_str()));

        // 요약된 텍스트가 포함된 문서를 기존 input.docx에 덮어쓰기
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/input.docx");

        cout << "요약된 내용을 input.docx에 저장 완료!" << endl;

    }
    catch (const System::Exception& e) {
        cerr << "Aspose 예외 발생: " << e->get_Message().ToUtf8String() << endl;
    }
    catch (const std::exception& e) {
        cerr << "std::exception 발생: " << e.what() << endl;
    }
    catch (...) {
        cerr << "알 수 없는 예외 발생" << endl;
    }
}