#include "Arduino.h"
#include "MotorControl.h"

Motor::Motor(int Dir1, int Dir2, int En1)
{
	pinMode(Dir1, OUTPUT);
	pinMode(Dir2, OUTPUT);
	_Dir1 = Dir1;
	_Dir2 = Dir2;
	_En1 = En1;
}

void Motor::Forward(int Pwr1)
{
	digitalWrite(_Dir1, HIGH);
	digitalWrite(_Dir2, LOW);
	analogWrite(_En1, Pwr1);
}
void Motor::Backword(int Pwr2)
{
	digitalWrite(_Dir1, LOW);
	digitalWrite(_Dir2, HIGH);
	analogWrite(_En1, Pwr2);

}

void Motor::Stop()
{
	digitalWrite(_Dir1, HIGH);
	digitalWrite(_Dir2, HIGH);
	analogWrite(_En1, 0);
}