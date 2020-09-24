// Header with blend mode functions
// Thanks to Jens Gruschel for the blend codes, great job !!!

// Defines to help with blend modes
#define BLM_AVERAGE			1			// Average mode
#define BLM_MULTIPLY		2			// Multiply mode
#define	BLM_SCREEN			3			// Screen mode
#define BLM_DARKEN			4			// Darken mode
#define BLM_LIGHTEN			5			// Lighten mode
#define BLM_DIFFERENCE		6			// Difference mode
#define BLM_NEGATION		7			// Negation mode
#define BLM_EXCLUSION		8			// Exclusion mode
#define	BLM_OVERLAY			9			// Overlay mode
#define BLM_HARDLIGHT		10			// Hard Light mode
#define	BLM_SOFTLIGHT		11			// Soft Light mode
#define BLM_COLORDODGE		12			// Color Dodge mode
#define BLM_COLORBURN		13			// Color Burn mode
#define BLM_SOFTDODGE		14			// Soft dodge mode
#define BLM_SOFTBURN		15			// Soft burn mode
#define BLM_REFLECT			16			// Reflect mode
#define BLM_GLOW			17			// Glow mode
#define BLM_FREEZE			18			// Freeze mode
#define BLM_HEAT			19			// Heat mode
#define BLM_ADDITIVE		20			// Additive mode
#define BLM_SUBTRACTIVE		21			// Subtractive mode
#define BLM_INTERPOLATION	22			// Interpolation mode
#define BLM_STAMP			23			// Stamp mode
#define BLM_XOR				24			// XOR mode

inline BYTE AverageMode (BYTE Color1, BYTE Color2);
inline BYTE MultiplyMode (BYTE Color1, BYTE Color2);
inline BYTE ScreenMode (BYTE Color1, BYTE Color2);
inline BYTE DarkenMode (BYTE Color1, BYTE Color2);
inline BYTE LightenMode (BYTE Color1, BYTE Color2);
inline BYTE DifferenceMode (BYTE Color1, BYTE Color2);
inline BYTE NegationMode (BYTE Color1, BYTE Color2);
inline BYTE ExclusionMode (BYTE Color1, BYTE Color2);
inline BYTE OverlayMode (BYTE Color1, BYTE Color2);
inline BYTE HardLightMode (BYTE Color1, BYTE Color2);
inline BYTE SoftLightMode (BYTE Color1, BYTE Color2);
inline BYTE ColorDodgeMode (BYTE Color1, BYTE Color2);
inline BYTE ColorBurnMode (BYTE Color1, BYTE Color2);
inline BYTE SoftDodgeMode (BYTE Color1, BYTE Color2);
inline BYTE SoftBurnMode (BYTE Color1, BYTE Color2);
inline BYTE ReflectMode (BYTE Color1, BYTE Color2);
inline BYTE GlowMode (BYTE Color1, BYTE Color2);
inline BYTE FreezeMode (BYTE Color1, BYTE Color2);
inline BYTE HeatMode (BYTE Color1, BYTE Color2);
inline BYTE AdditiveMode (BYTE Color1, BYTE Color2);
inline BYTE SubtractiveMode (BYTE Color1, BYTE Color2);
inline BYTE InterpolationMode (BYTE Color1, BYTE Color2);
inline BYTE StampMode (BYTE Color1, BYTE Color2);
inline BYTE XORMode (BYTE Color1, BYTE Color2);

BYTE ApplyBlendMode (BYTE Color1, BYTE Color2, int Mode);