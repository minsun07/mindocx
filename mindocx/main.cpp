#include <Aspose.Words.Cpp/Document.h>                                                         // word ������ ���� �� ó���� ���� Ŭ����
#include <Aspose.Words.Cpp/DocumentBuilder.h>                                                  // ������ �ý�Ʈ�� ��Ҹ� �߰��ϴ� Ŭ����
#include <Aspose.Words.Cpp/Drawing/Shape.h>                                                    // �׸�, ������ ���� �� ��ġ�� ���� Ŭ����
#include <Aspose.Words.Cpp/Drawing/RelativeHorizontalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/RelativeVerticalPosition.h> 
#include <Aspose.Words.Cpp/Drawing/WrapType.h>                                                 
#include <Aspose.Words.Cpp/Font.h>                                                              // ��Ʈ �������� ���� Ŭ����
#include <iostream>

using namespace Aspose::Words;                                                                 // ���� ���� �� ������ ���� Ŭ������ ����
using namespace Aspose::Words::Drawing;                                                        // ������ ���� �� �׷��� ��Ҹ� �ٷ�
using namespace System;                                                                        // Aspose ���̺귯������ ���Ǵ� �ֿ� ������ Ÿ��(SharedPtr ��)�� ����
using namespace std;


int main() {
    try {
        // ���� ���� �� ���� �ʱ�ȭ
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);   // ������ ����� ���� ����

        builder->get_Font()->set_Size(14);      // �ؽ�Ʈ ũ��
        builder->get_Font()->set_Bold(true);    // �ؽ�Ʈ ����
        builder->Writeln(u"Hello World!");      // �ؽ�Ʈ ���

        System::SharedPtr<Shape> shape = builder->InsertImage(  // �̹��� ���
            u"C:\Anne2.png",                                    // �̹��� ���
            RelativeHorizontalPosition::Page, 60.0,             // ����
            RelativeVerticalPosition::Page, 60.0,               // ����
            -1.0, -1.0,                                         // �̹��� ũ��
            WrapType::None                                      // �ؽ�Ʈ���� �� �ٲ� ���(�ؽ�Ʈ�� ��ġ��X)
        );

        // ���� ����
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/testWord1.docx");  //������ �ٲ㼭 �����ص� ���� ������
        cout << "���� ���� ����" << endl;

        

     /* ���� ����(2) : ������ ������ �� ����
     
        doc = System::MakeObject<Document>();
        builder = System::MakeObject<DocumentBuilder>(doc);
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/�׽�Ʈ����.docx");
        cout << "����2" << endl; */

    }catch (const System::Exception& e) {
        cerr << "Error: " << e->get_Message().ToUtf8String() << endl;
        
        // const : ���� ��ü�� ������� �ʵ��� ��.
        // & : ���� ��ü�� ������ ����. �������� �ʰ� ������ �����Ͽ� ������ ���
    }
}
