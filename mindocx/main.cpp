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
        // ���� ���� �� ���� �ʱ�ȭ
        System::SharedPtr<Document> doc = System::MakeObject<Document>();
        System::SharedPtr<DocumentBuilder> builder = System::MakeObject<DocumentBuilder>(doc);

        // ���� ����
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/�׽�Ʈ����.docx");  //������ �ٲ㼭 �����ص� ���� ������
        cout << "���� ���� ����" << endl;

  /*      doc = System::MakeObject<Document>();
        builder = System::MakeObject<DocumentBuilder>(doc);
        doc->Save(u"C:/2024-CPP-project/mindocx/x64/Debug/�׽�Ʈ����.docx"); // ���� ����(2) : �̷��� ������ ������ �� ����
        cout << "����2" << endl;*/


    }
    catch (const System::Exception& e) {
        cerr << "Error: " << e->get_Message().ToUtf8String() << endl;
        // const : ���� ��ü�� ������� �ʵ��� ��.
        // & : ���� ��ü�� �������� �ʰ� ������ �����Ͽ� ������ ���
        // e : ���� ��ü�� �̸�
        // cerr : console error stream�� ����. ǥ�� ���� ��� ��Ʈ��. 
        //        ���α׷� ���� �� �߻��� ���� �޽����� ����ϴµ� ���
        // e->get_Message() : e���� �߻��� ������ �޽����� �������� �޼���
        // .ToUtf8String() : ��ȯ�� �޽����� UTF-8�� ��ȯ
    }
}
