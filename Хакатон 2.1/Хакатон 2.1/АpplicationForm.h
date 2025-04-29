#pragma once
#include <msclr/marshal_cppstd.h>
#include "List.h"

namespace StarterForm {

	using namespace System;
	using namespace System::ComponentModel;
	using namespace System::Collections;
	using namespace System::Windows::Forms;
	using namespace System::Data;
	using namespace System::Drawing;

	/// <summary>
	/// ������ ��� �pplicationForm
	/// </summary>
	public ref class �pplicationForm : public System::Windows::Forms::Form
	{
	public:
		�pplicationForm(void)
		{
			InitializeComponent();
			//
			//TODO: �������� ��� ������������
			//
		}

	protected:
		/// <summary>
		/// ���������� ��� ������������ �������.
		/// </summary>
		~�pplicationForm()
		{
			if (components)
			{
				delete components;
			}
		}
	private: System::Windows::Forms::Button^ AddApl;



	private: System::Windows::Forms::Button^ DeletApl;

	private:
		System::Collections::Generic::List<Label^>^ appLabels = gcnew System::Collections::Generic::List<Label^>();
		System::Collections::Generic::List<Label^>^ typeLabels = gcnew System::Collections::Generic::List<Label^>();
		System::Collections::Generic::List<TextBox^>^ textBoxes = gcnew System::Collections::Generic::List<TextBox^>();
		System::Collections::Generic::List<Label^>^ LabelImage = gcnew System::Collections::Generic::List<Label^>();
		System::Collections::Generic::List<Button^>^ ButtonImage = gcnew System::Collections::Generic::List<Button^>();
		System::Collections::Generic::List<Label^>^ LabelInf = gcnew System::Collections::Generic::List<Label^>();

	protected:

	protected:

	private:
		/// <summary>
		/// ������������ ���������� ������������.
		/// </summary>
		System::ComponentModel::Container ^components;

#pragma region Windows Form Designer generated code
		/// <summary>
		/// ��������� ����� ��� ��������� ������������ � �� ��������� 
		/// ���������� ����� ������ � ������� ��������� ����.
		/// </summary>
		void InitializeComponent(void)
		{
			this->AddApl = (gcnew System::Windows::Forms::Button());
			this->DeletApl = (gcnew System::Windows::Forms::Button());
			this->SuspendLayout();
			// 
			// AddApl
			// 
			this->AddApl->Location = System::Drawing::Point(12, 12);
			this->AddApl->Name = L"AddApl";
			this->AddApl->Size = System::Drawing::Size(80, 39);
			this->AddApl->TabIndex = 0;
			this->AddApl->Text = L"�������� ����������";
			this->AddApl->UseVisualStyleBackColor = true;
			this->AddApl->Click += gcnew System::EventHandler(this, &�pplicationForm::AddApl_Click);
			// 
			// DeletApl
			// 
			this->DeletApl->Location = System::Drawing::Point(209, 12);
			this->DeletApl->Name = L"DeletApl";
			this->DeletApl->Size = System::Drawing::Size(80, 39);
			this->DeletApl->TabIndex = 4;
			this->DeletApl->Text = L"������ ����������";
			this->DeletApl->UseVisualStyleBackColor = true;
			this->DeletApl->Click += gcnew System::EventHandler(this, &�pplicationForm::DeletApl_Click);
			// 
			// �pplicationForm
			// 
			this->AutoScaleDimensions = System::Drawing::SizeF(6, 13);
			this->AutoScaleMode = System::Windows::Forms::AutoScaleMode::Font;
			this->ClientSize = System::Drawing::Size(301, 380);
			this->Controls->Add(this->DeletApl);
			this->Controls->Add(this->AddApl);
			this->Name = L"�pplicationForm";
			this->Text = L"�pplicationForm";
			this->ResumeLayout(false);

		}
#pragma endregion
	int count = 0;
	int currentY = 60;

	private: System::Void AddApl_Click(System::Object^ sender, System::EventArgs^ e) {
		std::string arr[28] = {"�", "�", "�", "�", "�", "�"};

		if (count < 6) {
			Label^ labelAppName = gcnew Label();
			std::string NameStr = "���������� " + arr[count];
			System::String^ Name = msclr::interop::marshal_as<System::String^>(NameStr);
			labelAppName->Text = Name;
			labelAppName->Location = System::Drawing::Point(14, currentY);
			labelAppName->Size = System::Drawing::Size(150, 19);
			labelAppName->Font = gcnew System::Drawing::Font("Times New Roman", 12, System::Drawing::FontStyle::Bold);
			this->Controls->Add(labelAppName);
			appLabels->Add(labelAppName);

			sharedData::ListStorage::imagePaths->Add(gcnew System::Collections::Generic::List<String^>());
			sharedData::ListStorage::appTypes->Add("");

			currentY += 19;

			Label^ labelTypeName = gcnew Label();
			labelTypeName->Text = "��� ����������";
			labelTypeName->Location = System::Drawing::Point(14, currentY);
			labelTypeName->Size = System::Drawing::Size(131, 19);
			labelTypeName->Font = gcnew System::Drawing::Font("Times New Roman", 12, System::Drawing::FontStyle::Bold);
			this->Controls->Add(labelTypeName);
			typeLabels->Add(labelTypeName);

			TextBox^ textBoxAppName = gcnew TextBox();
			textBoxAppName->Location = System::Drawing::Point(152, currentY);
			textBoxAppName->Multiline = true;
			textBoxAppName->Size = System::Drawing::Size(137, 46);
			textBoxAppName->Font = gcnew System::Drawing::Font("Times New Roman", 12);
			textBoxAppName->Tag = count;
			textBoxAppName->TextChanged += gcnew System::EventHandler(this, &�pplicationForm::TextBoxAppName_TextChanged);
			this->Controls->Add(textBoxAppName);
			textBoxes->Add(textBoxAppName);

			currentY += 52;

			Label^ labelImage = gcnew Label();
			labelImage->Text = "�������� �����������";
			labelImage->Location = System::Drawing::Point(14, currentY);
			labelImage->Size = System::Drawing::Size(179, 19);
			labelImage->Font = gcnew System::Drawing::Font("Times New Roman", 12, System::Drawing::FontStyle::Bold);
			this->Controls->Add(labelImage);
			LabelImage->Add(labelImage);

			Button^ buttonImage = gcnew Button();
			buttonImage->Text = "�������";
			buttonImage->Location = System::Drawing::Point(199, currentY);
			buttonImage->Size = System::Drawing::Size(90, 46);
			buttonImage->Font = gcnew System::Drawing::Font("Times New Roman", 12);
			buttonImage->Click += gcnew System::EventHandler(this, &�pplicationForm::ImageBut_Click);
			buttonImage->Tag = count;
			this->Controls->Add(buttonImage);
			ButtonImage->Add(buttonImage);

			currentY += 19;

			Label^ labelInf = gcnew Label();
			labelInf->Text = "";
			labelInf->Location = System::Drawing::Point(14, currentY);
			labelInf->Size = System::Drawing::Size(179, 19);
			labelInf->Font = gcnew System::Drawing::Font("Times New Roman", 12, System::Drawing::FontStyle::Bold);
			labelInf->ForeColor = System::Drawing::Color::ForestGreen;
			this->Controls->Add(labelInf);
			LabelInf->Add(labelInf);

			currentY += 70;
			count++;
		}
		else {
			MessageBox::Show("��������� ������������ ���������� ����������!", "������", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: System::Void DeletApl_Click(System::Object^ sender, System::EventArgs^ e) {
		if (count > 0)
		{
			int idx = count - 1;  // ������ ���������� "����������"

			// 1. ������� �������� �� �����
			this->Controls->Remove(appLabels[idx]);
			this->Controls->Remove(typeLabels[idx]);
			this->Controls->Remove(textBoxes[idx]);
			this->Controls->Remove(LabelImage[idx]);
			this->Controls->Remove(ButtonImage[idx]);
			this->Controls->Remove(LabelInf[idx]);

			// 2. ������� ���� ������� �� �������
			appLabels->RemoveAt(idx);
			typeLabels->RemoveAt(idx);
			textBoxes->RemoveAt(idx);
			LabelImage->RemoveAt(idx);
			ButtonImage->RemoveAt(idx);
			LabelInf->RemoveAt(idx);

			// 3. ������ sharedData
			sharedData::ListStorage::imagePaths->RemoveAt(idx);
			sharedData::ListStorage::appTypes->RemoveAt(idx);

			// 4. ����������� ������� � "��������" ����������� Y
			count--;
			currentY -= 160; // ��� �� ������� �� � ���� ���������� ��� ����������
		}
		else
		{
			MessageBox::Show("��� ���������� ��� ��������.", "������", MessageBoxButtons::OK, MessageBoxIcon::Error);
		}
	}
	private: int countImg = 0;
	private: System::Void ImageBut_Click(System::Object^ sender, System::EventArgs^ e) {

		Button^ clickedButton = dynamic_cast<Button^>(sender);
		int appIndex = safe_cast<int>(clickedButton->Tag);

		OpenFileDialog^ fileSelect = gcnew OpenFileDialog();

		fileSelect->Filter = "����������� (*.png, *.jpg)|*.png;*.jpg";
		fileSelect->Title = "�������� �����������";

		if (fileSelect->ShowDialog() == System::Windows::Forms::DialogResult::OK)
		{
			String^ selectedPath = fileSelect->FileName;
			sharedData::ListStorage::imagePaths[appIndex]->Add(selectedPath);

			Label^ infoLabel = LabelInf[appIndex];
			infoLabel->Text = "��������� " + sharedData::ListStorage::imagePaths[appIndex]->Count.ToString();
		}
	}
	private: System::Void TextBoxAppName_TextChanged(System::Object^ sender, System::EventArgs^ e) {
		TextBox^ changedTextBox = dynamic_cast<TextBox^>(sender);
		int appIndex = safe_cast<int>(changedTextBox->Tag);

		sharedData::ListStorage::appTypes[appIndex] = changedTextBox->Text;
	}
};
}
