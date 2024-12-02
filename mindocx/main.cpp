#include <Aspose.Words.Cpp/Document.h>
#include <Aspose.Words.Cpp/DocumentBuilder.h>
#include <Aspose.Words.Cpp/Drawing/Shape.h>
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h>
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h>
#include <Aspose.Words.Cpp/Drawing/WrapType.h>
#include <iostream>

using namespace Aspose::Words;
using namespace Aspose::Words::Drawing;
using namespace System;
using namespace std;

int main() {
    try {
        // 문서 생성 및 빌더 초기화
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // 문서 저장
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/테스트문서.docx");  //문서명만 바꿔서 실행해도 문서 생성됨
        cout << "문서 저장 성공" << endl;

  /*      doc = System::MakeObject<Document>();
        builder = System::MakeObject<DocumentBuilder>(doc);
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/테스트문서.docx"); // 문서 저장(2) : 이렇게 문서를 생성할 수 있음
        cout << "성공2" << endl;*/


    }
    catch (const System::Exception& e) {
        cerr << "Error: " << e->get_Message().ToUtf8String() << endl;
        // const : 예외 객체가 변경되지 않도록 함.
        // & : 예외 객체를 복사하지 않고 원본을 참조하여 성능을 향상
        // e : 예외 객체의 이름
        // cerr : console error stream의 약자. 표준 에러 출력 스트림. 
        //        프로그램 실행 중 발생한 에러 메시지를 출력하는데 사용
        // e->get_Message() : e에서 발생한 예외의 메시지를 가져오는 메서드
        // .ToUtf8String() : 반환된 메시지를 UTF-8로 변환
    }
}
