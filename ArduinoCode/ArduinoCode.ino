#include "MotorControl.h"
int ans;
Motor Motor(4,5,6);
void setup()
{
	Serial.begin(9600);
}
void loop()
{
	//ans = 0;
	if (Serial.available()>0)
	{
		ans = Serial.parseInt();
		if (ans == 1)
		{
			Motor.Forward(255);
			delay(500);
			Motor.Stop();
		}
		else
		{
			Motor.Backword(200);
			delay(500);
			Motor.Stop();
		}
	}
}