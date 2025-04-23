#pragma once
#include <cstdlib>
#include <msclr/marshal_cppstd.h>
#include "TitlePageData.h"
#include "TitlePage.h"

namespace StarterForm {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// Сводка для StarterForm
	/// </summary>
	public ref class StarterForm : public System::Windows::Forms::Form
	{

	private:
		TitlePageData* data;
	private: System::Windows::Forms::Button^ SavePdf;

		   TitlePage^ titlePage;

	public:
		StarterForm(void)
		{
			InitializeComponent();
			data = new TitlePageData();
			titlePage = gcnew TitlePage(data);
		}

	protected:
		/// <summary>
		/// Освободить все используемые ресурсы.
		/// </summary>
		~StarterForm()
		{
			if (components)
			{
				delete components;
			}
			if (data) delete data;
		}
	private: System::Windows::Forms::Label^ label1;
	private: System::DirectoryServices::DirectorySearcher^ directorySearcher1;
	private: System::DirectoryServices::DirectoryEntry^ directoryEntry1;
	private: System::Windows::Forms::Button^ Select;
	private: System::Windows::Forms::Button^ ConfirmDocx;
	private: System::Windows::Forms::Button^ SelectSave;

	private: System::Windows::Forms::Label^ label2;
	private: System::Windows::Forms::Button^ SelectSame;
	private: System::Windows::Forms::Label^ error;
	private: System::Windows::Forms::Button^ TitlePageLabel;

	private: System::Windows::Forms::Button^ Switch;





	protected:

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
			this->label1 = (gcnew System::Windows::Forms::Label());
			this->directorySearcher1 = (gcnew System::DirectoryServices::DirectorySearcher());
			this->directoryEntry1 = (gcnew System::DirectoryServices::DirectoryEntry());
			this->Select = (gcnew System::Windows::Forms::Button());
			this->ConfirmDocx = (gcnew System::Windows::Forms::Button());
			this->SelectSave = (gcnew System::Windows::Forms::Button());
			this->label2 = (gcnew System::Windows::Forms::Label());
			this->SelectSame = (gcnew System::Windows::Forms::Button());
			this->error = (gcnew System::Windows::Forms::Label());
			this->TitlePageLabel = (gcnew System::Windows::Forms::Button());
			this->Switch = (gcnew System::Windows::Forms::Button());
			this->SavePdf = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// label1
			// 
			this->label1->AutoSize = true;
			this->label1->Font = (gcnew System::Drawing::Font(L"a_CooperBlack", 14.25F, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label1->Location = System::Drawing::Point(12, 9);
			this->label1->Name = L"label1";
			this->label1->Size = System::Drawing::Size(277, 21);
			this->label1->TabIndex = 0;
			this->label1->Text = L"Выберите путь до файла";
			// 
			// directorySearcher1
			// 
			this->directorySearcher1->ClientTimeout = System::TimeSpan::Parse(L"-00:00:01");
			this->directorySearcher1->ServerPageTimeLimit = System::TimeSpan::Parse(L"-00:00:01");
			this->directorySearcher1->ServerTimeLimit = System::TimeSpan::Parse(L"-00:00:01");
			// 
			// Select
			// 
			this->Select->Location = System::Drawing::Point(16, 33);
			this->Select->Name = L"Select";
			this->Select->Size = System::Drawing::Size(273, 33);
			this->Select->TabIndex = 1;
			this->Select->Text = L"Указать путь";
			this->Select->UseVisualStyleBackColor = true;
			this->Select->Click += gcnew System::EventHandler(this, &StarterForm::Select_Click);
			// 
			// ConfirmDocx
			// 
			this->ConfirmDocx->Location = System::Drawing::Point(92, 350);
			this->ConfirmDocx->Name = L"ConfirmDocx";
			this->ConfirmDocx->Size = System::Drawing::Size(124, 49);
			this->ConfirmDocx->TabIndex = 2;
			this->ConfirmDocx->Text = L"Отформатировать docx";
			this->ConfirmDocx->UseVisualStyleBackColor = true;
			this->ConfirmDocx->Click += gcnew System::EventHandler(this, &StarterForm::ConfirmDocx_Click);
			// 
			// SelectSave
			// 
			this->SelectSave->Location = System::Drawing::Point(16, 116);
			this->SelectSave->Name = L"SelectSave";
			this->SelectSave->Size = System::Drawing::Size(134, 57);
			this->SelectSave->TabIndex = 4;
			this->SelectSave->Text = L"Указать путь";
			this->SelectSave->UseVisualStyleBackColor = true;
			this->SelectSave->Click += gcnew System::EventHandler(this, &StarterForm::SelectSave_Click);
			// 
			// label2
			// 
			this->label2->AutoSize = true;
			this->label2->Font = (gcnew System::Drawing::Font(L"a_CooperBlack", 12, System::Drawing::FontStyle::Regular, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->label2->Location = System::Drawing::Point(12, 95);
			this->label2->Name = L"label2";
			this->label2->Size = System::Drawing::Size(291, 18);
			this->label2->TabIndex = 3;
			this->label2->Text = L"Выберите путь для сохранения";
			// 
			// SelectSame
			// 
			this->SelectSame->Location = System::Drawing::Point(155, 116);
			this->SelectSame->Name = L"SelectSame";
			this->SelectSame->Size = System::Drawing::Size(134, 57);
			this->SelectSame->TabIndex = 5;
			this->SelectSame->Text = L"Сохранить туда же где и файл";
			this->SelectSame->UseVisualStyleBackColor = true;
			this->SelectSame->Click += gcnew System::EventHandler(this, &StarterForm::SelectSame_Click);
			// 
			// error
			// 
			this->error->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 11.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->error->ForeColor = System::Drawing::Color::DarkRed;
			this->error->Location = System::Drawing::Point(15, 402);
			this->error->Name = L"error";
			this->error->Size = System::Drawing::Size(274, 23);
			this->error->TabIndex = 6;
			this->error->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
			// 
			// TitlePageLabel
			// 
			this->TitlePageLabel->Enabled = false;
			this->TitlePageLabel->Location = System::Drawing::Point(155, 216);
			this->TitlePageLabel->Name = L"TitlePageLabel";
			this->TitlePageLabel->Size = System::Drawing::Size(134, 57);
			this->TitlePageLabel->TabIndex = 7;
			this->TitlePageLabel->Text = L"Настроить титульный лист";
			this->TitlePageLabel->UseVisualStyleBackColor = true;
			this->TitlePageLabel->Click += gcnew System::EventHandler(this, &StarterForm::TitlePageLabel_Click);
			// 
			// Switch
			// 
			this->Switch->Location = System::Drawing::Point(46, 216);
			this->Switch->Name = L"Switch";
			this->Switch->Size = System::Drawing::Size(70, 57);
			this->Switch->TabIndex = 8;
			this->Switch->Text = L"Добавить титульный лист";
			this->Switch->UseVisualStyleBackColor = true;
			this->Switch->Click += gcnew System::EventHandler(this, &StarterForm::Switch_Click);
			// 
			// SavePdf
			// 
			this->SavePdf->Location = System::Drawing::Point(92, 295);
			this->SavePdf->Name = L"SavePdf";
			this->SavePdf->Size = System::Drawing::Size(124, 49);
			this->SavePdf->TabIndex = 9;
			this->SavePdf->Text = L"Конвертировать docx в pdf";
			this->SavePdf->UseVisualStyleBackColor = true;
			this->SavePdf->Click += gcnew System::EventHandler(this, &StarterForm::SavePdf_Click);
			// 
			// StarterForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(305, 432);
			this->Controls->Add(this->SavePdf);
			this->Controls->Add(this->Switch);
			this->Controls->Add(this->TitlePageLabel);
			this->Controls->Add(this->error);
			this->Controls->Add(this->SelectSame);
			this->Controls->Add(this->SelectSave);
			this->Controls->Add(this->label2);
			this->Controls->Add(this->ConfirmDocx);
			this->Controls->Add(this->Select);
			this->Controls->Add(this->label1);
			this->Name = L"StarterForm";
			this->Text = L"StarterForm";
			this->ResumeLayout(false);
			this->PerformLayout();

		}
#pragma endregion
	private: System::Void Select_Click(System::Object^ sender, System::EventArgs^ e) {
		OpenFileDialog^ fileSelect = gcnew OpenFileDialog();

		fileSelect->Filter = "Word документы (*.docx)|*.docx";
		fileSelect->Title = "Выберите документ Word";

		if (fileSelect->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			Select->Text = fileSelect->FileName;
		}
	}

	private: System::Void ConfirmDocx_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ inputPath = Select->Text;
		String^ outputPath = SelectSave->Text + ".docx";

		if (outputPath == "Указать путь")
		{
			error->Text = "Выберите путь для сохранения";
		}

		if (inputPath == "Указать путь")
		{
			error->Text = "Выберите путь до файла";
		}

		if (inputPath != "Указать путь" && outputPath != "Указать путь")
		{
			if (TitlePageLabel->Enabled == true)
			{
				bool usePdf = false;
				String^ command = "python \"..\\..\\DocxFormat\\DocxFormat\\DocxFormat.py\" \""
					+ inputPath + "\" \"" + outputPath + "\" \"" + usePdf + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Inst) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Depart) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Proj) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Doc) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Discip) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Numb) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Spec) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->FIOStud) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->Group) + "\""
					+ " \"" + msclr::interop::marshal_as<String^>(data->FIOVis) + "\"";

				std::string arguments = msclr::interop::marshal_as<std::string>(command);

				system(arguments.c_str());

			}
			else
			{
				bool usePdf = false;
				String^ command = "python \"..\\..\\DocxFormat\\DocxFormat\\DocxFormat.py\" \"" + inputPath + "\" \"" + outputPath + "\" \"" + usePdf + "\"";
				std::string arguments = msclr::interop::marshal_as<std::string>(command);
				system(arguments.c_str());
			}
		}
	}

	private: System::Void SelectSave_Click(System::Object^ sender, System::EventArgs^ e) {
		FolderBrowserDialog^ folderSelect = gcnew FolderBrowserDialog();

		folderSelect->Description = "Выберите папку для сохранения документа";

		if (folderSelect->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			SelectSave->Text = folderSelect->SelectedPath + "\\Результат";
		}
	}

	private: System::Void SelectSame_Click(System::Object^ sender, System::EventArgs^ e) {
		if (Select->Text == "Указать путь")
		{
			error->Text = "Выберите путь до файла";
		}
		else 
		{
			String^ folderPath = System::IO::Path::GetDirectoryName(Select->Text) + "Результат";
			SelectSave->Text = folderPath;
		}
	}
	private: System::Void TitlePageLabel_Click(System::Object^ sender, System::EventArgs^ e) {
		titlePage->ShowDialog();
	}
	private: System::Void Switch_Click(System::Object^ sender, System::EventArgs^ e) {
		if (TitlePageLabel->Enabled == false)
		{
			Switch->Text = "Убрать титульный лист";
			TitlePageLabel->Enabled = true;
		}
		else 
		{
			Switch->Text = "Добавить титульный лист";
			TitlePageLabel->Enabled = false;
		}
	}
	private: System::Void SavePdf_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ inputPath = Select->Text;
		String^ outputPath = SelectSave->Text + ".docx";

		if (outputPath == "Указать путь")
		{
			error->Text = "Выберите путь для сохранения";
		}

		if (inputPath == "Указать путь")
		{
			error->Text = "Выберите путь до файла";
		}

		if (inputPath != "Указать путь" && outputPath != "Указать путь")
		{
			bool usePdf = true;
			String^ command = "python \"..\\..\\DocxFormat\\DocxFormat\\DocxFormat.py\" \"" + inputPath + "\" \"" + outputPath + "\" \"" + usePdf + "\"";
			std::string arguments = msclr::interop::marshal_as<std::string>(command);
			system(arguments.c_str());
		}
	}
};
}
