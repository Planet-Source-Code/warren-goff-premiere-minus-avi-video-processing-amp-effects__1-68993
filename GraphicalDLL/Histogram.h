// Header with functions to analize the histogram table
#include "GlobalFunctions.h"

// Defines to be used with histogram functions
#define HST_RED				1			// Red
#define HST_GREEN			2			// Green
#define HST_BLUE			4			// Blue
#define HST_COLOR			7			// All the colors
#define HST_GRAY			8			// Gray

#define GRAD_METALLIC		1			// Metallic
#define GRAD_GOLD			2			// Gold gradient
#define GRAD_ICE			3			// Ice gradient

// Struct to be used with histogram functions
typedef struct
{
	int Table[256];
	int Minimum;
	int Maximum;
} Histogram;

// Struct to store all the amplitude values of a table
typedef struct
{
	int Low;
	int High;
	int LowRed;
	int LowGreen;
	int LowBlue;
	int HighRed;
	int HighGreen;
	int HighBlue;
} ColorAmp;

// Functions to get or set histogram information
void FindHistoMinAndMaxValues (Histogram *Histo);

void GetHistogram (unsigned char	Bits[],
				   int				Width,
				   int				Height,
				   int				Histo[],
				   int				Flag);

void AssignTables (UCHAR	*RedTable, 
				   UCHAR	*GreenTable, 
				   UCHAR	*BlueTable, 
				   UCHAR	Bits[],  
				   int		Width, 
				   int		Height);

void ApplyMetallicLayer (UCHAR	Bits[],
						 int	Width,
						 int	Height,
						 int	Levels);

void ApplyMetallicShiftLayer (UCHAR	Bits[],
							  int	Width,
							  int	Height,
							  int	Levels,
							  int	Shift);

void ApplyGoldLayer (UCHAR	Bits[],
					 int	Width,
					 int	Height);

void MakeGradient (ColorAmp *cAmp,
				   UCHAR	*rTable,
				   UCHAR	*gTable,
				   UCHAR	*bTable);

void ShiftTable (UCHAR	*Table,
				 int	Shift);

void ApplyIceLayer (UCHAR	Bits[],
					int		Width,
					int		Height);
