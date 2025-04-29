#pragma once
#include <cstdlib>
#include <msclr/marshal_cppstd.h>
#include "TitlePageData.h"
#include "TitlePage.h"
#include "АpplicationForm.h"
#include "List.h"

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
	private: System::Windows::Forms::Button^ Switch1;

	private: System::Windows::Forms::Button^ applicationBut;


		   TitlePage^ titlePage;
	private: System::Windows::Forms::Button^ CheckErrorBut;


		   АpplicationForm^ applicationForm;

	public:
		StarterForm(void)
		{
			InitializeComponent();
			data = new TitlePageData();
			titlePage = gcnew TitlePage(data);
			applicationForm = gcnew АpplicationForm();
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
			this->Switch1 = (gcnew System::Windows::Forms::Button());
			this->applicationBut = (gcnew System::Windows::Forms::Button());
			this->CheckErrorBut = (gcnew System::Windows::Forms::Button());
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
			this->ConfirmDocx->Location = System::Drawing::Point(92, 404);
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
			this->SelectSame->Text = L"Сохранить в ту же директорию";
			this->SelectSame->UseVisualStyleBackColor = true;
			this->SelectSame->Click += gcnew System::EventHandler(this, &StarterForm::SelectSame_Click);
			// 
			// error
			// 
			this->error->Font = (gcnew System::Drawing::Font(L"Microsoft Sans Serif", 11.25F, System::Drawing::FontStyle::Bold, System::Drawing::GraphicsUnit::Point,
				static_cast<System::Byte>(204)));
			this->error->ForeColor = System::Drawing::Color::DarkRed;
			this->error->Location = System::Drawing::Point(15, 516);
			this->error->Name = L"error";
			this->error->Size = System::Drawing::Size(274, 23);
			this->error->TabIndex = 6;
			this->error->TextAlign = System::Drawing::ContentAlignment::MiddleCenter;
			// 
			// TitlePageLabel
			// 
			this->TitlePageLabel->Enabled = false;
			this->TitlePageLabel->Location = System::Drawing::Point(155, 196);
			this->TitlePageLabel->Name = L"TitlePageLabel";
			this->TitlePageLabel->Size = System::Drawing::Size(134, 57);
			this->TitlePageLabel->TabIndex = 7;
			this->TitlePageLabel->Text = L"Настроить титульный лист";
			this->TitlePageLabel->UseVisualStyleBackColor = true;
			this->TitlePageLabel->Click += gcnew System::EventHandler(this, &StarterForm::TitlePageLabel_Click);
			// 
			// Switch
			// 
			this->Switch->Location = System::Drawing::Point(29, 196);
			this->Switch->Name = L"Switch";
			this->Switch->Size = System::Drawing::Size(96, 57);
			this->Switch->TabIndex = 8;
			this->Switch->Text = L"Добавить титульный лист и содержание";
			this->Switch->UseVisualStyleBackColor = true;
			this->Switch->Click += gcnew System::EventHandler(this, &StarterForm::Switch_Click);
			// 
			// SavePdf
			// 
			this->SavePdf->Location = System::Drawing::Point(92, 349);
			this->SavePdf->Name = L"SavePdf";
			this->SavePdf->Size = System::Drawing::Size(124, 49);
			this->SavePdf->TabIndex = 9;
			this->SavePdf->Text = L"Конвертировать docx в pdf";
			this->SavePdf->UseVisualStyleBackColor = true;
			this->SavePdf->Click += gcnew System::EventHandler(this, &StarterForm::SavePdf_Click);
			// 
			// Switch1
			// 
			this->Switch1->Location = System::Drawing::Point(29, 266);
			this->Switch1->Name = L"Switch1";
			this->Switch1->Size = System::Drawing::Size(96, 57);
			this->Switch1->TabIndex = 10;
			this->Switch1->Text = L"Добавить приложение";
			this->Switch1->UseVisualStyleBackColor = true;
			this->Switch1->Click += gcnew System::EventHandler(this, &StarterForm::Switch1_Click);
			// 
			// applicationBut
			// 
			this->applicationBut->Enabled = false;
			this->applicationBut->Location = System::Drawing::Point(155, 259);
			this->applicationBut->Name = L"applicationBut";
			this->applicationBut->Size = System::Drawing::Size(134, 57);
			this->applicationBut->TabIndex = 11;
			this->applicationBut->Text = L"Настроить приложение";
			this->applicationBut->UseVisualStyleBackColor = true;
			this->applicationBut->Click += gcnew System::EventHandler(this, &StarterForm::applicationBut_Click);
			// 
			// CheckErrorBut
			// 
			this->CheckErrorBut->Location = System::Drawing::Point(92, 459);
			this->CheckErrorBut->Name = L"CheckErrorBut";
			this->CheckErrorBut->Size = System::Drawing::Size(124, 49);
			this->CheckErrorBut->TabIndex = 12;
			this->CheckErrorBut->Text = L"Проверить docx на ошибки";
			this->CheckErrorBut->UseVisualStyleBackColor = true;
			this->CheckErrorBut->Click += gcnew System::EventHandler(this, &StarterForm::CheckErrorBut_Click);
			// 
			// StarterForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(311, 551);
			this->Controls->Add(this->CheckErrorBut);
			this->Controls->Add(this->applicationBut);
			this->Controls->Add(this->Switch1);
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

	private:
		std::string replaceSpaces(std::string str) {
			for (size_t i = 0; i < str.size(); ++i) {
				if (str[i] == ' ') {
					str[i] = '_';
				}
			}
			return str;
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
			bool usePdf = false;
			std::vector<std::string> cmdArgs;
			cmdArgs.push_back("python \"..\\..\\DocxFormat\\DocxFormat\\DocxFormat.py\" ");
			cmdArgs.push_back(msclr::interop::marshal_as<std::string>(inputPath));
			cmdArgs.push_back(msclr::interop::marshal_as<std::string>(outputPath));
			if (usePdf) cmdArgs.push_back("--pdf");

			if (TitlePageLabel->Enabled) {
				cmdArgs.push_back("--title");
				cmdArgs.push_back(replaceSpaces(data->Inst));
				cmdArgs.push_back(replaceSpaces(data->Depart));
				cmdArgs.push_back(replaceSpaces(data->Proj));
				cmdArgs.push_back(replaceSpaces(data->Doc));
				cmdArgs.push_back(replaceSpaces(data->Discip));
				cmdArgs.push_back(replaceSpaces(data->Numb));
				cmdArgs.push_back(replaceSpaces(data->Spec));
				cmdArgs.push_back(replaceSpaces(data->FIOStud));
				cmdArgs.push_back(replaceSpaces(data->Group));
				cmdArgs.push_back(replaceSpaces(data->FIOVis));

				cmdArgs.push_back(replaceSpaces(data->InstRed));
				cmdArgs.push_back(replaceSpaces(data->DocType));
				cmdArgs.push_back(replaceSpaces(data->UDC));
				cmdArgs.push_back(replaceSpaces(data->SubDoc));
			}

			if (applicationBut->Enabled) {
				int appsCount = sharedData::ListStorage::appTypes->Count;
				cmdArgs.push_back("--appdata");
				cmdArgs.push_back(std::to_string(appsCount));
				for (int i = 0; i < appsCount; ++i) {

					cmdArgs.push_back(msclr::interop::marshal_as<std::string>(sharedData::ListStorage::appTypes[i]));

					int imgCount = sharedData::ListStorage::imagePaths[i]->Count;
					cmdArgs.push_back(std::to_string(imgCount));

					auto images = sharedData::ListStorage::imagePaths[i];
					for (int j = 0; j < imgCount; ++j)
					{
						cmdArgs.push_back(msclr::interop::marshal_as<std::string>(images[j]));
					}
				}
			}

			std::ostringstream oss;
			for (auto& a : cmdArgs) {
				oss << a << " ";
			}

			auto cmdLine = oss.str();

			system(cmdLine.c_str());
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
			String^ folderPath = System::IO::Path::GetDirectoryName(Select->Text) + "\\Результат";
			SelectSave->Text = folderPath;
		}
	}
	private: System::Void TitlePageLabel_Click(System::Object^ sender, System::EventArgs^ e) {
		titlePage->ShowDialog();
	}
	private: System::Void Switch_Click(System::Object^ sender, System::EventArgs^ e) {
		if (TitlePageLabel->Enabled == false)
		{
			Switch->Text = "Убрать титульный лист и содержание";
			TitlePageLabel->Enabled = true;
		}
		else 
		{
			Switch->Text = "Добавить титульный лист и содержание";
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
			std::vector<std::string> cmdArgs;
			cmdArgs.push_back("python \"..\\..\\DocxFormat\\DocxFormat\\DocxFormat.py\" ");
			cmdArgs.push_back(msclr::interop::marshal_as<std::string>(inputPath));
			cmdArgs.push_back(msclr::interop::marshal_as<std::string>(outputPath));
			if (usePdf) cmdArgs.push_back("--pdf");

			std::ostringstream oss;
			for (auto& a : cmdArgs) {
				oss << a << " ";
			}

			auto cmdLine = oss.str();

			system(cmdLine.c_str());
		}
	}
	private: System::Void Switch1_Click(System::Object^ sender, System::EventArgs^ e) {
		if (applicationBut->Enabled == false)
		{
			Switch1->Text = "Убрать приложение";
			applicationBut->Enabled = true;
		}
		else
		{
			Switch1->Text = "Добавить приложение";
			applicationBut->Enabled = false;
		}
	}
	private: System::Void applicationBut_Click(System::Object^ sender, System::EventArgs^ e) {
		applicationForm->ShowDialog();
	}
	private: System::Void CheckErrorBut_Click(System::Object^ sender, System::EventArgs^ e) {
		String^ inputPath = Select->Text;
		String^ outputPath = SelectSave->Text + "_Ошибки.docx";

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
			std::vector<std::string> cmdArgs;
			cmdArgs.push_back("python \"..\\..\\DocxFormat\\DocxFormat\\DocxFormat.py\" ");
			cmdArgs.push_back(msclr::interop::marshal_as<std::string>(inputPath));
			cmdArgs.push_back(msclr::interop::marshal_as<std::string>(outputPath));
			cmdArgs.push_back("--check");

			std::ostringstream oss;
			for (auto& a : cmdArgs) {
				oss << a << " ";
			}

			auto cmdLine = oss.str();

			system(cmdLine.c_str());
		}
	}
};
}
