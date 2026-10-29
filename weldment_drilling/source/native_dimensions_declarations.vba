' Pure VBA dimensions, ported from autosizedrilling Geometry/AutoSizeDrill.
Private Const ND_PI As Double = 3.14159265358979
Private Const ND_MODEL_TOL As Double = 0.00005
Private Const ND_DIAMETER_TOL As Double = 0.000001
Private Type NDLine
    X1 As Double
    Y1 As Double
    X2 As Double
    Y2 As Double
    Entity As SldWorks.Edge
End Type
Private Type NDHole
    X As Double
    Y As Double
    Radius As Double
    Entity As SldWorks.Edge
End Type
Private Type NDBounds
    MinX As Double
    MinY As Double
    MaxX As Double
    MaxY As Double
    LeftIndex As Long
    RightIndex As Long
    BottomIndex As Long
    TopIndex As Long
End Type
