#include "frmMain.h"
#include "ddcmonitors.h"

using namespace System;
using namespace System::Windows::Forms;
using namespace PureLogic;

[STAThread]
int WINAPI WinMain(HINSTANCE inst, HINSTANCE prev, LPSTR cmd, int show)
{
	Application::EnableVisualStyles();
	Application::SetCompatibleTextRenderingDefault(false);

	PureLogic::frmMain form;
	Application::Run(%form);

	return 0;
}


void frmMain::frmMain_Load(System::Object^  sender, System::EventArgs^  e) {
	//MessageBox::Show("Startup test");
}

void frmMain::frmMain_Paint(System::Object^  sender, PaintEventArgs^  e) {
	//System::Drawing::Graphics ^g = e->Graphics;
	
	//Drawing::SolidBrush ^blockBack = gcnew Drawing::SolidBrush(Color::FromArgb(255, 222, 140));
	//Drawing::SolidBrush ^programBackground = gcnew Drawing::SolidBrush(Color::FromArgb(180, 180, 180));
	//Drawing::Pen ^blockBorder = gcnew Pen(Drawing::Color::Black);
	

	//paint background
	//e->Graphics->FillRectangle(programBackground, frmMain::ClientRectangle);

	
	

	//draw background of block
	//e->Graphics->FillRectangle(blockBack, 40, 40, 25, 35);
	
	//e->Graphics->DrawRectangle(blockBorder, 40, 40, 25, 35);
}

Void frmMain::button1_Click(System::Object^  sender, System::EventArgs^  e) {
	ddcmonitors bla;
	//bla.SetPower(ddcmonitors::Power::On);
	bla.SetBackLight((int)(this->numericUpDown1->Value));
}
