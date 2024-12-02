#include <Aspose.Words.Cpp/Document.h>                                                         // word 문서의 생성 및 처리를 위한 클래스
#include <Aspose.Words.Cpp/DocumentBuilder.h>                                                  // 문서에 택스트나 요소를 추가하는 클래스
#include <Aspose.Words.Cpp/Drawing/Shape.h>                                                    // 그림, 도형을 삽입 및 위치를 위한 클래스
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/WrapType.h>                                                 
#include <Aspose.Words.Cpp/Font.h>                                                              // 폰트 디자인을 위한 클래스
#include <iostream>

using namespace Aspose::Words;                                                                 // 문서 생성 및 조작을 위한 클래스가 포함
using namespace Aspose::Words::Drawing;                                                        // 문서에 도형 및 그래픽 요소를 다룸
using namespace System;                                                                        // Aspose 라이브러리에서 사용되는 주요 데이터 타입(SharedPtr 등)을 포함
using namespace std;


int main() {
    try {
        // 문서 생성 및 빌더 초기화
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);   // 문서와 연결된 빌더 생성

        builder->get_Font()->set_Size(14);      // 텍스트 크기
        builder->get_Font()->set_Bold(true);    // 텍스트 굵기
        builder->Writeln(u"Hello World!");      // 텍스트 출력

        System::SharedPtr<Shape> shape = builder->InsertImage(  // 이미지 출력
            u"C:\Anne2.png",                                    // 이미지 경로
            RelativeHorizontalPosition::Page, 60.0,             // 수평
            RelativeVerticalPosition::Page, 60.0,               // 수직
            -1.0, -1.0,                                         // 이미지 크기
            WrapType::None                                      // 텍스트와의 줄 바꿈 방식(텍스트와 겹치지X)
        );

        // 문서 저장
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/testWord1.docx");  //문서명만 바꿔서 실행해도 문서 생성됨
        cout << "문서 저장 성공" << endl;

        

     /* 문서 저장(2) : 문서를 생성할 수 있음
     
        doc = System::MakeObject<Document>();
        builder = System::MakeObject<DocumentBuilder>(doc);
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/테스트문서.docx");
        cout << "성공2" << endl; */

    }catch (const System::Exception& e) {
        cerr << "Error: " << e->get_Message().ToUtf8String() << endl;
        
        // const : 예외 객체가 변경되지 않도록 함.
        // & : 예외 객체를 참조로 전달. 복사하지 않고 원본을 참조하여 성능을 향상
    }
}
