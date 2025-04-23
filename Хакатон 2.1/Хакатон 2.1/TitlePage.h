#pragma once
#include <msclr/marshal_cppstd.h>
#include "fstream"
#include <string>
#include <codecvt>
#include <sstream>
#include <locale>
#include "json.hpp"
#include "TitlePageData.h"
using json = nlohmann::json;

namespace StarterForm {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Сводка для TitlePage
	/// </summary>
	public ref class TitlePage : public System::Windows::Forms::Form
	{

	private: TitlePageData* data;

	public:
		TitlePage(TitlePageData* data_)
		{
			InitializeComponent();
			data = data_;
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~TitlePage()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Label^ InstLab;
	private: System::Windows::Forms::TextBox^ InstBox;
	private: System::Windows::Forms::TextBox^ DepBox;
	protected:

	protected:

	protected:


	private: System::Windows::Forms::Label^ DepLab;
	private: System::Windows::Forms::TextBox^ DisBox;


	private: System::Windows::Forms::Label^ DisLab;

	private: System::Windows::Forms::Label^ ProjLab;
	private: System::Windows::Forms::TextBox^ DocBox;



	private: System::Windows::Forms::Label^ DocLab;
	private: System::Windows::Forms::TextBox^ NumbBox;

	private: System::Windows::Forms::Label^ NumbLab;
	private: System::Windows::Forms::TextBox^ SpecBox;





	private: System::Windows::Forms::Label^ SpecLab;

	private: System::Windows::Forms::TextBox^ ProjBox;
	private: System::Windows::Forms::TextBox^ FIOStudBox;

	private: System::Windows::Forms::Label^ FIOStudLab;
	private: System::Windows::Forms::TextBox^ FIOVisBox;





	private: System::Windows::Forms::Label^ FIOVisLab;

	private: System::Windows::Forms::TextBox^ GroupBox;

	private: System::Windows::Forms::Label^ GroupLab;

	private: System::Windows::Forms::Button^ LoadBut;

	private: System::Windows::Forms::Button^ SaveBut;
	private: System::Windows::Forms::Button^ ConfirmBut;
	private: System::Windows::Forms::Label^ InfLabel;



	private:
		/// <summary>
		/// Обязательная переменная конструктора.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// Требуемый метод для поддержки конструктора — не изменяйте 
		/// содержимое этого метода с помощью редактора кода.
		/// </summary>
		void InitializeComponent(void)
		{
			this->InstLab = (gcnew System::Windows::Forms::Label());
			this->InstBox = (gcnew System::Windows::Forms::TextBox());
			this->DepBox = (gcnew System::Windows::Forms::TextBox());
			this->DepLab = (gcnew System::Windows::Forms::Label());
			this->DisBox = (gcnew System::Windows::Forms::TextBox());
			this->DisLab = (gcnew System::Windows::Forms::Label());
			this->ProjLab = (gcnew System::Windows::Forms::Label());
			this->DocBox = (gcnew System::Windows::Forms::TextBox());
			this->DocLab = (gcnew System::Windows::Forms::Label());
			this->NumbBox = (gcnew System::Windows::Forms::TextBox());
			this->NumbLab = (gcnew System::Windows::Forms::Label());
			this->SpecBox = (gcnew System::Windows::Forms::TextBox());
			this->SpecLab = (gcnew System::Windows::Forms::Label());
			this->ProjBox = (gcnew System::Windows::Forms::TextBox());
			this->FIOStudBox = (gcnew System::Windows::Forms::TextBox());
			this->FIOStudLab = (gcnew System::Windows::Forms::Label());
			this->FIOVisBox = (gcnew System::Windows::Forms::TextBox());
			this->FIOVisLab = (gcnew System::Windows::Forms::Label());
			this->GroupBox = (gcnew System::Windows::Forms::TextBox());
			this->GroupLab = (gcnew System::Windows::Forms::Label());
			this->LoadBut = (gcnew System::Windows::Forms::Button());
			this->SaveBut = (gcnew System::Windows::Forms::Button());
			this->ConfirmBut = (gcnew System::Windows::Forms::Button());
			this->InfLabel = (gcnew System::Windows::Forms::Label());
			this->SuspendLayout();
			// 
			// InstLab
			// 
			this->InstLab->AutoSize = true;
			this->InstLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->InstLab->Location = System::Drawing::Point(12, 12);
			this->InstLab->Name = L"InstLab";
			this->InstLab->Size = System::Drawing::Size(81, 16);
			this->InstLab->TabIndex = 0;
			this->InstLab->Text = L"Институт:";
			// 
			// InstBox
			// 
			this->InstBox->Location = System::Drawing::Point(99, 8);
			this->InstBox->Multiline = true;
			this->InstBox->Name = L"InstBox";
			this->InstBox->Size = System::Drawing::Size(273, 23);
			this->InstBox->TabIndex = 1;
			// 
			// DepBox
			// 
			this->DepBox->Location = System::Drawing::Point(15, 60);
			this->DepBox->Multiline = true;
			this->DepBox->Name = L"DepBox";
			this->DepBox->Size = System::Drawing::Size(354, 23);
			this->DepBox->TabIndex = 3;
			// 
			// DepLab
			// 
			this->DepLab->AutoSize = true;
			this->DepLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->DepLab->Location = System::Drawing::Point(12, 41);
			this->DepLab->Name = L"DepLab";
			this->DepLab->Size = System::Drawing::Size(77, 16);
			this->DepLab->TabIndex = 2;
			this->DepLab->Text = L"Кафедра:";
			// 
			// DisBox
			// 
			this->DisBox->Location = System::Drawing::Point(15, 214);
			this->DisBox->Multiline = true;
			this->DisBox->Name = L"DisBox";
			this->DisBox->Size = System::Drawing::Size(357, 23);
			this->DisBox->TabIndex = 5;
			// 
			// DisLab
			// 
			this->DisLab->AutoSize = true;
			this->DisLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->DisLab->Location = System::Drawing::Point(12, 195);
			this->DisLab->Name = L"DisLab";
			this->DisLab->Size = System::Drawing::Size(101, 16);
			this->DisLab->TabIndex = 4;
			this->DisLab->Text = L"Дисциплина:";
			// 
			// ProjLab
			// 
			this->ProjLab->AutoSize = true;
			this->ProjLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->ProjLab->Location = System::Drawing::Point(12, 89);
			this->ProjLab->Name = L"ProjLab";
			this->ProjLab->Size = System::Drawing::Size(150, 16);
			this->ProjLab->TabIndex = 6;
			this->ProjLab->Text = L"Название проекта:";
			// 
			// DocBox
			// 
			this->DocBox->Location = System::Drawing::Point(15, 160);
			this->DocBox->Multiline = true;
			this->DocBox->Name = L"DocBox";
			this->DocBox->Size = System::Drawing::Size(354, 23);
			this->DocBox->TabIndex = 9;
			// 
			// DocLab
			// 
			this->DocLab->AutoSize = true;
			this->DocLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->DocLab->Location = System::Drawing::Point(12, 141);
			this->DocLab->Name = L"DocLab";
			this->DocLab->Size = System::Drawing::Size(123, 16);
			this->DocLab->TabIndex = 8;
			this->DocLab->Text = L"Вид документа:";
			// 
			// NumbBox
			// 
			this->NumbBox->Location = System::Drawing::Point(194, 252);
			this->NumbBox->Multiline = true;
			this->NumbBox->Name = L"NumbBox";
			this->NumbBox->Size = System::Drawing::Size(178, 23);
			this->NumbBox->TabIndex = 11;
			// 
			// NumbLab
			// 
			this->NumbLab->AutoSize = true;
			this->NumbLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->NumbLab->Location = System::Drawing::Point(12, 256);
			this->NumbLab->Name = L"NumbLab";
			this->NumbLab->Size = System::Drawing::Size(176, 16);
			this->NumbLab->TabIndex = 10;
			this->NumbLab->Text = L"Номер специальности:";
			// 
			// SpecBox
			// 
			this->SpecBox->Location = System::Drawing::Point(10, 304);
			this->SpecBox->Multiline = true;
			this->SpecBox->Name = L"SpecBox";
			this->SpecBox->Size = System::Drawing::Size(362, 23);
			this->SpecBox->TabIndex = 13;
			// 
			// SpecLab
			// 
			this->SpecLab->AutoSize = true;
			this->SpecLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->SpecLab->Location = System::Drawing::Point(12, 285);
			this->SpecLab->Name = L"SpecLab";
			this->SpecLab->Size = System::Drawing::Size(202, 16);
			this->SpecLab->TabIndex = 12;
			this->SpecLab->Text = L"Название специальности:";
			// 
			// ProjBox
			// 
			this->ProjBox->Location = System::Drawing::Point(15, 108);
			this->ProjBox->Multiline = true;
			this->ProjBox->Name = L"ProjBox";
			this->ProjBox->Size = System::Drawing::Size(357, 23);
			this->ProjBox->TabIndex = 14;
			// 
			// FIOStudBox
			// 
			this->FIOStudBox->Location = System::Drawing::Point(136, 333);
			this->FIOStudBox->Multiline = true;
			this->FIOStudBox->Name = L"FIOStudBox";
			this->FIOStudBox->Size = System::Drawing::Size(236, 23);
			this->FIOStudBox->TabIndex = 16;
			// 
			// FIOStudLab
			// 
			this->FIOStudLab->AutoSize = true;
			this->FIOStudLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->FIOStudLab->Location = System::Drawing::Point(12, 337);
			this->FIOStudLab->Name = L"FIOStudLab";
			this->FIOStudLab->Size = System::Drawing::Size(118, 16);
			this->FIOStudLab->TabIndex = 15;
			this->FIOStudLab->Text = L"ФИО студента:";
			// 
			// FIOVisBox
			// 
			this->FIOVisBox->Location = System::Drawing::Point(175, 393);
			this->FIOVisBox->Multiline = true;
			this->FIOVisBox->Name = L"FIOVisBox";
			this->FIOVisBox->Size = System::Drawing::Size(197, 23);
			this->FIOVisBox->TabIndex = 18;
			// 
			// FIOVisLab
			// 
			this->FIOVisLab->AutoSize = true;
			this->FIOVisLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->FIOVisLab->Location = System::Drawing::Point(15, 397);
			this->FIOVisLab->Name = L"FIOVisLab";
			this->FIOVisLab->Size = System::Drawing::Size(154, 16);
			this->FIOVisLab->TabIndex = 17;
			this->FIOVisLab->Text = L"ФИО руководителя:";
			// 
			// GroupBox
			// 
			this->GroupBox->Location = System::Drawing::Point(136, 362);
			this->GroupBox->Multiline = true;
			this->GroupBox->Name = L"GroupBox";
			this->GroupBox->Size = System::Drawing::Size(236, 23);
			this->GroupBox->TabIndex = 20;
			// 
			// GroupLab
			// 
			this->GroupLab->AutoSize = true;
			this->GroupLab->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 9.75F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->GroupLab->Location = System::Drawing::Point(14, 366);
			this->GroupLab->Name = L"GroupLab";
			this->GroupLab->Size = System::Drawing::Size(116, 16);
			this->GroupLab->TabIndex = 19;
			this->GroupLab->Text = L"Номер группы:";
			// 
			// LoadBut
			// 
			this->LoadBut->Location = System::Drawing::Point(12, 441);
			this->LoadBut->Name = L"LoadBut";
			this->LoadBut->Size = System::Drawing::Size(176, 38);
			this->LoadBut->TabIndex = 21;
			this->LoadBut->Text = L"Загрузить шалбон";
			this->LoadBut->UseVisualStyleBackColor = true;
			this->LoadBut->Click += gcnew System::EventHandler(this, &TitlePage::LoadBut_Click);
			// 
			// SaveBut
			// 
			this->SaveBut->Location = System::Drawing::Point(193, 441);
			this->SaveBut->Name = L"SaveBut";
			this->SaveBut->Size = System::Drawing::Size(176, 38);
			this->SaveBut->TabIndex = 22;
			this->SaveBut->Text = L"Сохранить шаблон";
			this->SaveBut->UseVisualStyleBackColor = true;
			this->SaveBut->Click += gcnew System::EventHandler(this, &TitlePage::SaveBut_Click);
			// 
			// ConfirmBut
			// 
			this->ConfirmBut->Location = System::Drawing::Point(103, 517);
			this->ConfirmBut->Name = L"ConfirmBut";
			this->ConfirmBut->Size = System::Drawing::Size(176, 38);
			this->ConfirmBut->TabIndex = 23;
			this->ConfirmBut->Text = L"ОК";
			this->ConfirmBut->UseVisualStyleBackColor = true;
			this->ConfirmBut->Click += gcnew System::EventHandler(this, &TitlePage::ConfirmBut_Click);
			// 
			// InfLabel
			// 
			this->InfLabel->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 11.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->InfLabel->ForeColor = System::Drawing::Color::DarkRed;
			this->InfLabel->Location = System::Drawing::Point(15, 491);
			this->InfLabel->Name = L"InfLabel";
			this->InfLabel->Size = System::Drawing::Size(354, 23);
			this->InfLabel->TabIndex = 24;
			this->InfLabel->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
			// 
			// TitlePage
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(384, 560);
			this->Controls->Add(this->InfLabel);
			this->Controls->Add(this->ConfirmBut);
			this->Controls->Add(this->SaveBut);
			this->Controls->Add(this->LoadBut);
			this->Controls->Add(this->GroupBox);
			this->Controls->Add(this->GroupLab);
			this->Controls->Add(this->FIOVisBox);
			this->Controls->Add(this->FIOVisLab);
			this->Controls->Add(this->FIOStudBox);
			this->Controls->Add(this->FIOStudLab);
			this->Controls->Add(this->ProjBox);
			this->Controls->Add(this->SpecBox);
			this->Controls->Add(this->SpecLab);
			this->Controls->Add(this->NumbBox);
			this->Controls->Add(this->NumbLab);
			this->Controls->Add(this->DocBox);
			this->Controls->Add(this->DocLab);
			this->Controls->Add(this->ProjLab);
			this->Controls->Add(this->DisBox);
			this->Controls->Add(this->DisLab);
			this->Controls->Add(this->DepBox);
			this->Controls->Add(this->DepLab);
			this->Controls->Add(this->InstBox);
			this->Controls->Add(this->InstLab);
			this->Name = L"TitlePage";
			this->Text = L"TitlePage";
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void SaveBut_Click(System::Object^ sender, System::EventArgs^ e) {

		FolderBrowserDialog^ folderSelect = gcnew FolderBrowserDialog();

		folderSelect->Description = "Выберите папку для сохранения шаблона";

		if (folderSelect->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			std::string Path = msclr::interop::marshal_as<std::string>(folderSelect->SelectedPath + "\\TitlePagePattern.json");

			std::wofstream file(Path);
			file.imbue(std::locale(std::locale(), new std::codecvt_utf8<wchar_t>));

			if (file.is_open()) {
				json j;
				msclr::interop::marshal_context ctx;
				std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> converter;

				j["institute"] = converter.to_bytes(ctx.marshal_as<std::wstring>(InstBox->Text));
				j["department"] = converter.to_bytes(ctx.marshal_as<std::wstring>(DepBox->Text));
				j["project"] = converter.to_bytes(ctx.marshal_as<std::wstring>(ProjBox->Text));
				j["doc_type"] = converter.to_bytes(ctx.marshal_as<std::wstring>(DocBox->Text));
				j["discipline"] = converter.to_bytes(ctx.marshal_as<std::wstring>(DisBox->Text));
				j["number"] = converter.to_bytes(ctx.marshal_as<std::wstring>(NumbBox->Text));
				j["specialty"] = converter.to_bytes(ctx.marshal_as<std::wstring>(SpecBox->Text));
				j["student"] = converter.to_bytes(ctx.marshal_as<std::wstring>(FIOStudBox->Text));
				j["group"] = converter.to_bytes(ctx.marshal_as<std::wstring>(GroupBox->Text));
				j["teacher"] = converter.to_bytes(ctx.marshal_as<std::wstring>(FIOVisBox->Text));

				file << converter.from_bytes(j.dump(4));
				file.close();

				InfLabel->ForeColor = System::Drawing::Color::Green;
				InfLabel->Text = "Шаблон успешно сохранён";
			}
			else {
				InfLabel->ForeColor = System::Drawing::Color::DarkRed;
				InfLabel->Text = "Шаблон не сохранён";
			}
		}
		else {
			InfLabel->ForeColor = System::Drawing::Color::DarkRed;
			InfLabel->Text = "Шаблон не сохранён";
		}
	}
	private: System::Void LoadBut_Click(System::Object^ sender, System::EventArgs^ e) {

		OpenFileDialog^ fileSelect = gcnew OpenFileDialog();

		fileSelect->Filter = "JSON файл (*.json)|*.json";
		fileSelect->Title = "Выберите шаблон";

		if (fileSelect->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			std::string Path = msclr::interop::marshal_as<std::string>(fileSelect->FileName);

			std::ifstream file(Path, std::ios::binary);

			std::ostringstream buf;
			buf << file.rdbuf();
			file.close();
			std::string content = buf.str();

			json j = json::parse(content);

			std::wstring_convert<std::codecvt_utf8_utf16<wchar_t>> conv;

			InstBox->Text = gcnew System::String(conv.from_bytes(j["institute"].get<std::string>()).c_str());
			DepBox->Text = gcnew System::String(conv.from_bytes(j["department"].get<std::string>()).c_str());
			ProjBox->Text = gcnew System::String(conv.from_bytes(j["project"].get<std::string>()).c_str());
			DocBox->Text = gcnew System::String(conv.from_bytes(j["doc_type"].get<std::string>()).c_str());
			DisBox->Text = gcnew System::String(conv.from_bytes(j["discipline"].get<std::string>()).c_str());
			NumbBox->Text = gcnew System::String(conv.from_bytes(j["number"].get<std::string>()).c_str());
			SpecBox->Text = gcnew System::String(conv.from_bytes(j["specialty"].get<std::string>()).c_str());
			FIOStudBox->Text = gcnew System::String(conv.from_bytes(j["student"].get<std::string>()).c_str());
			GroupBox->Text = gcnew System::String(conv.from_bytes(j["group"].get<std::string>()).c_str());
			FIOVisBox->Text = gcnew System::String(conv.from_bytes(j["teacher"].get<std::string>()).c_str());

			InfLabel->ForeColor = System::Drawing::Color::Green;
			InfLabel->Text = "Шаблон успешно загружен";
		}
		else {
			InfLabel->ForeColor = System::Drawing::Color::DarkRed;
			InfLabel->Text = "Шаблон не загружен";
		}
	}
	private: System::Void ConfirmBut_Click(System::Object^ sender, System::EventArgs^ e) {

		data->Inst = msclr::interop::marshal_as<std::string>(InstBox->Text);
		data->Depart = msclr::interop::marshal_as<std::string>(DepBox->Text);
		data->Proj = msclr::interop::marshal_as<std::string>(ProjBox->Text);
		data->Doc = msclr::interop::marshal_as<std::string>(DocBox->Text);
		data->Discip = msclr::interop::marshal_as<std::string>(DisBox->Text);
		data->Numb = msclr::interop::marshal_as<std::string>(NumbBox->Text);
		data->Spec = msclr::interop::marshal_as<std::string>(SpecBox->Text);
		data->FIOStud = msclr::interop::marshal_as<std::string>(FIOStudBox->Text);
		data->Group = msclr::interop::marshal_as<std::string>(GroupBox->Text);
		data->FIOVis = msclr::interop::marshal_as<std::string>(FIOVisBox->Text);

		this->Hide();
	}
};
}
