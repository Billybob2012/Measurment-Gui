#ifndef MotorControl_h
#include "Arduino.h"

class Motor
{
public:
	Motor(int Dir1, int Dir2, int En1);
	void Forward(int Pwr1);
	void Backword(int Pwr2);
	void Stop();
private:
	int	_Dir1;
	int _Dir2;
	int _En1;
};
#endif