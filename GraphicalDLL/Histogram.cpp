#include "Histogram.h"

/* ProportionalValue function														*
 *																					*
 * DestColor		=> Destiny colour												*
 * SrcColor			=> Source colour												*
 * Shade			=> Value to find the intermediate								*
 *																					*
 * Theory			=> This function does the same thing that ShadeColors function	*
 *					but using double variables.										*
 *																					*/
inline double GradientValue (double FirstValue, 
							 double SecondValue, 
							 double Gradient)
{
	if (Gradient == 0.0)			// if shade is 0 return the DestColor
		return FirstValue;
	if (Gradient == 255.0)		// if shade is 255 return the SrcColor
		return SecondValue;
	// Now, calcule the intermediate colour
	return ((FirstValue * (255 - Gradient) + SecondValue * Gradient) / 256);
}

/* Function to return the histogram table											*
 *																					*
 * Bits[]			=> Image bit array												*
 * Width			=> Image width													*
 * Height			=> Image height													*
 * Histo			=> Table to store the histogram									*
 * Flag				=> Indicates what histogram should be stored					*														*
 *																					*
 * Theory			=> With image bits, I'm trying to get the histogram. Count each	*
 *					color in that picture.											*
 *																					*/
void GetHistogram (unsigned char	Bits[],
				   int				Width,
				   int				Height,
				   int				Histo[],
				   int				Flag)
{
	int i, j, index;
	int LineWidth = 3 * Width;
	if (LineWidth % 4)						// Don't take off this step
		LineWidth += (4 - LineWidth % 4);

	for (i = 0; i < 256; i++)
		Histo[i] = 0;

	for (i = 0; i < Width; i++)
		for (j = 0; j < Height; j++)
		{
			index = j * LineWidth + 3 * i;
			if (Flag == HST_RED)					// Count Red values
				Histo[Bits[index+2]]++;
			else if (Flag == HST_GREEN)				// Count Green values
				Histo[Bits[index+1]]++;
			else if (Flag == HST_BLUE)				// Count Blue values
				Histo[Bits[ index ]]++;
		}
	
	return;
}

/* Function to find the minimum and maximum values in the histogram					*
 *																					*
 * *Histo			=> Histogram table to analize									*
 *																					*
 * Theory			=> With histogram table, we find the first and the last color	*
 *																					*/
void FindHistoMinAndMaxValues (Histogram *Histo)
{
	int Min = 0, Max = 255;

	while ((Histo->Table[Min] == 0) && (Min < 255))
		Min++;
	while ((Histo->Table[Max] == 0) && (Max > 0))
		Max--;
	
	Histo->Minimum = Min;
	Histo->Maximum = Max;
}

void ApplyMetallicLayer (UCHAR	Bits[],
						 int	Width,
						 int	Height,
						 int	Levels)
{
	int j, k = 0;
	UCHAR *mTable = new UCHAR[COLOR_SIZE];

	if (Levels < 2)
		return;

	for (j = 0; j < 255; )
	{
		for (k = 0; k < 256; k += Levels)
			mTable[j++] = (UCHAR)k;
		for (k = 255; k > 0; k -= Levels)
			mTable[j++] = (UCHAR)k;
	}
	mTable[255] = (Levels % 2 == 0) ? 0 : 255;

	AssignTables (mTable, mTable, mTable, Bits, Width, Height);
	delete [] mTable;
	return;
}

void ApplyMetallicShiftLayer (UCHAR	Bits[],
							  int	Width,
							  int	Height,
							  int	Levels,
							  int	Shift)
{
	int i, factor = 255 / Levels;
	ColorAmp cAmp;
	UCHAR *mTable = new UCHAR[COLOR_SIZE];

	if (Levels < 1)
		return;

	for (i = 0; i < 256; i++)
		mTable[i] = 0;

	for (i = 0; i < Levels; i++)
	{
		if (i % 2)
		{
			cAmp.Low = i * factor;
			cAmp.LowRed = 255;
			cAmp.LowGreen = 255;
			cAmp.LowBlue = 255;
			cAmp.High = (i + 1) * factor;
			cAmp.HighRed = 0;
			cAmp.HighGreen = 0;
			cAmp.HighBlue = 0;
			mTable[255] = 0;
		}
		else
		{
			cAmp.Low = i * factor + 1;
			cAmp.LowRed = 0;
			cAmp.LowGreen = 0;
			cAmp.LowBlue = 0;
			cAmp.High = (i + 1) * factor;
			cAmp.HighRed = 255;
			cAmp.HighGreen = 255;
			cAmp.HighBlue = 255;
			mTable[255] = 255;
		}

		MakeGradient (&cAmp, mTable, mTable, mTable);
	}

	ShiftTable (mTable, Shift);
	AssignTables (mTable, mTable, mTable, Bits, Width, Height);
	delete [] mTable;
	return;
}

void MakeGradient (ColorAmp *cAmp,
				   UCHAR	*rTable,
				   UCHAR	*gTable,
				   UCHAR	*bTable)
{
	int i;
	double delta, temp;

	if (cAmp->High == cAmp->Low)
		return;

	delta = 255.0 / (cAmp->High - cAmp->Low);

	for (i = cAmp->Low; i <= cAmp->High; i++)
	{
		temp = (i - cAmp->Low) * delta;
		rTable[i] = (UCHAR)GradientValue (cAmp->LowRed,   cAmp->HighRed,   temp);
		gTable[i] = (UCHAR)GradientValue (cAmp->LowGreen, cAmp->HighGreen, temp);
		bTable[i] = (UCHAR)GradientValue (cAmp->LowBlue,  cAmp->HighBlue,  temp);
	}

	return;
}

void ShiftTable (UCHAR	*Table,
				 int	Shift)
{
	UCHAR *tempTable = new UCHAR[COLOR_SIZE];
	int NewPosition;

	memcpy (tempTable, Table, COLOR_SIZE);

	for (int i = 0; i < 256; i++)
	{
		NewPosition = abs(i + Shift) & 0x000000FF;
		Table[NewPosition] = tempTable[i];
	}

	delete [] tempTable;

	return;
}

void ApplyGoldLayer (UCHAR	Bits[],
					 int	Width,
					 int	Height)
{
	ColorAmp cAmp;
	UCHAR *rTable = new UCHAR[COLOR_SIZE]; 
	UCHAR *gTable = new UCHAR[COLOR_SIZE];
	UCHAR *bTable = new UCHAR[COLOR_SIZE];

	for (int i = 0; i < 256; i++)
		rTable[i] = gTable[i] = bTable[i] = 0;
	
	cAmp.Low = 0;
	cAmp.LowRed = 0;
	cAmp.LowGreen = 0;
	cAmp.LowBlue = 0;
	cAmp.High = 55;
	cAmp.HighRed = 190;
	cAmp.HighGreen = 55;
	cAmp.HighBlue = 0;
	MakeGradient (&cAmp, rTable, gTable, bTable);

	cAmp.Low = 55;
	cAmp.LowRed = 190;
	cAmp.LowGreen = 55;
	cAmp.LowBlue = 0;
	cAmp.High = 155;
	cAmp.HighRed = 255;
	cAmp.HighGreen = 190;
	cAmp.HighBlue = 50;
	MakeGradient (&cAmp, rTable, gTable, bTable);

	cAmp.Low = 155;
	cAmp.LowRed = 255;
	cAmp.LowGreen = 190;
	cAmp.LowBlue = 50;
	cAmp.High = 255;
	cAmp.HighRed = 255;
	cAmp.HighGreen = 255;
	cAmp.HighBlue = 255;
	MakeGradient (&cAmp, rTable, gTable, bTable);

	AssignTables (rTable, gTable, bTable, Bits, Width, Height);
	delete [] rTable;
	delete [] gTable;
	delete [] bTable;

	return;
}

void ApplyIceLayer (UCHAR	Bits[],
					int		Width,
					int		Height)
{
	ColorAmp cAmp;
	UCHAR *rTable = new UCHAR[COLOR_SIZE]; 
	UCHAR *gTable = new UCHAR[COLOR_SIZE];
	UCHAR *bTable = new UCHAR[COLOR_SIZE];

	for (int i = 0; i < 256; i++)
		rTable[i] = gTable[i] = bTable[i] = 0;
	
	cAmp.Low = 0;
	cAmp.LowRed = 0;
	cAmp.LowGreen = 0;
	cAmp.LowBlue = 0;
	cAmp.High = 55;
	cAmp.HighRed = 0;
	cAmp.HighGreen = 65;
	cAmp.HighBlue = 205;
	MakeGradient (&cAmp, rTable, gTable, bTable);

	cAmp.Low = 55;
	cAmp.LowRed = 0;
	cAmp.LowGreen = 65;
	cAmp.LowBlue = 205;
	cAmp.High = 155;
	cAmp.HighRed = 65;
	cAmp.HighGreen = 205;
	cAmp.HighBlue = 255;
	MakeGradient (&cAmp, rTable, gTable, bTable);

	cAmp.Low = 155;
	cAmp.LowRed = 65;
	cAmp.LowGreen = 205;
	cAmp.LowBlue = 255;
	cAmp.High = 255;
	cAmp.HighRed = 255;
	cAmp.HighGreen = 255;
	cAmp.HighBlue = 255;
	MakeGradient (&cAmp, rTable, gTable, bTable);

	AssignTables (rTable, gTable, bTable, Bits, Width, Height);
	delete [] rTable;
	delete [] gTable;
	delete [] bTable;

	return;
}
