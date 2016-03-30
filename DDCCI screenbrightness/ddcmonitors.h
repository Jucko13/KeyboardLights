#pragma once
#include <windows.h>
#include <set>


class ddcmonitors
{
public:
	enum Power { On = 0x01, Standby = 0x02, Suspend = 0x03, Off = 0x04, HardOff = 0x05 };

	ddcmonitors();
	virtual ~ddcmonitors();

	int GetAmount();
	void SetPower(Power nPower);
	void SetBackLight(int brightness);

private:
	static BOOL CALLBACK MonitorEnumProc(HMONITOR hMonitor, HDC hdcMonitor, LPRECT lprcMonitor, LPARAM dwData);

	std::set<HANDLE> MonitorHandles;
};

