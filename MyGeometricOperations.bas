Attribute VB_Name = "MyGeometricOperations"
' MyGeometricOperations
'----------------------------------------------------------------------------------
'  Jeff Jenness
'  Jenness Enterprises
'  3020 N. Schevene Blvd.
'  Flagstaff, AZ  86004
'  http://www.jennessent.com
'  jeffj@jennessent.com
'  phone:  1-928-607-4638
'----------------------------------------------------------------------------------
'                 ArcCosJen - Returns ArcCos (Inverse Cosine)
'                 ArcSinJen - Returns ArcSin (Inverse Sine)
'                 AsDegrees - CONVERTS RADIANS TO DEGREES
'                 AsRadians - CONVERTS DEGREES TO RADIANS
'                     atan2 - Given DeltaY and DeltaX, returns ArcTangent that is sensitive to quadrant.
'          AzimuthHaversine - GIVEN TWO GEOGRAPHIC POINTS, RETURNS STARTING AZIMUTH POINTING IN GREAT CIRCLE OVER SPHERE.
'          BufferGeographic - ESTIMATE BUFFER AROUND GEOGRAPHIC GEOMETRY, BY PROJECTING THAT GEOMETRY INTO AN AZIMUTHAL EQUIDISTANT
'                             PROJECTION CENTERED ON GEOMETRY ENVELOPE CENTROID.  AN ALTERNATIVE METHOD TO EstimateDistanceOnSphere
'               CalcBearing - GIVEN TWO POINTS, CALCULATES THE CARTESIAN BEARING, WHERE 0 = NORTH, 360 DEGREES GOING CLOCKWISE
'              CalcBearing2 - SAME AS CalcBearing, EXCEPT IT RETURNS -999 IF THE TWO POINTS ARE COINCIDENT.
'CalcDirectionDeviationDegrees - GIVES THE DIFFERENCE IN DEGREES BETWEEN ANGLE 1 AND ANGLE 2.  POSITIVE IF ANGLE 2 IS CLOCKWISE
'                             FROM ANGLE 1.
'        CalcBearingNumbers - SAME AS CalcBearing2, EXCEPT IT TAKES DOUBLE VALUES INSTEAD OF POINTS
'        CalcCheckClockwise - CHECKS IF 3 CONSECUTIVE POINTS ARE ARRANGED COUNTERCLOCKWISE
' CalcCheckClockwiseNumbers - IDENTICAL TO CalcCheckClockwise, BUT USING DOUBLES INSTEAD OF IPOINTS
'         CalcClosestPoints - GIVEN TWO GEOMETRIES AND OPTIONAL NUMBER OF TIMES TO GO BACK AND FORTH BETWEEN CURVES, RETURNS
'                             IArray CONTAINING:
'                             IStringArray containing either "Intersecting Shapes" or
'                                "Empty Shapes" + two booleans indicating which geometry is empty.
'                             -- OR --
'                             3 OBJECTS:
'                             0) Connector Line AS IPOLYLINE
'                             1) Closest Point on Geometry #1         AS IPOINT
'                             2) Closest Point on Geometry #2         AS IPOINT
'     CalcProjectedDistance - Given two projected points, returns distance using Pythagorean Theorem.
' CalcProjectedDistanceNumbers - Given two projected sets of coordinates, returns distance using Pythagorean Theorem.
'            CalcDistMatrix - GIVEN AN IARRAY OF SHAPES, RETURNS A COLLECTION WHERE:
'                             INDEX = IArrayIndex1 & "_" & IArrayIndex2, and Object =
'                             IArray of {Distance, Optional Line, Optional Azimuth}
'        CalcFarthestPoints - Given a Geometry, Method (Trig vs. Spherical vs. Spheroidal; geometry must be unprojected if
'                             Spherical or Spheroidal), and Placeholders for First Point (IPoint), Last Point (IPoint),
'                             Distance (Double), Starting Azimuth (Double), Ending Azimuth (Double), Starting Reverse Azimuth (Double),
'                             Ending Reverse Azimuth (Double):  Returns boolean stating whether it worked or not.
'CalcFarthestPointsByNumbers - Same as CalcFarthestPoints, but uses a double array of X- and Y-coordinates so it works faster.
'         CalcInternalAngle - GIVEN 3 POINTS, RETURNS THE INTERNAL ANGLE (IN DEGREES) AND OPTIONALLY THE ANGLE OF DEVIATION
'                             NOTE: ASSUMES PLANE, USES STANDARD TRIGONOMETRY
'             CalcPointLine - GIVEN POINT, DISTANCE, AZIMUTH, EMPTY ENDPOINT AND EMPTY POLYLINE, REPLACES EMPTY
'                             ENDPOINT WITH ACTUAL ENDPOINT AND OPTIONALLY RETURNS A POLYLINE CONNECTOR
'          CalcPointNumbers - LIKE CalcPointLine ABOVE, BUT WORKS WITH NUMBERS SO IS FASTER.  DOES NOT RETURN A POLYLINE, THOUGH.
'      CheckPointInTriangle - GIVEN COORDINATES FOR TRIANGLE VERTICES AND COORDINATES FOR ADDITIONAL POINT, RETURNS BOOLEAN
'                             STATING WHETHER POINT IS INSIDE TRIANGLE OR NOT
' ConvertAngleCompassDegreesToMathRadians - GIVEN A DIRECTION IN COMPASS DEGREES (WHICH STARTS AT NORTH AND GOES CLOCKWISE),
'                             RETURNS MATHEMATICAL DIRECTION (WHICH STARTS AT EAST AND GOES CLOCKWISE)
' ConvertRotationMathRadiansToCompassDegrees - GIVEN A MATHEMATICAL DIRECTION (WHICH STARTS AT EAST AND GOES CLOCKWISE),
'                             RETURNS DIRECTION IN COMPASS DEGREES (WHICH STARTS AT NORTH AND GOES CLOCKWISE)
' ConvertRotationCompassDegreesToMathRadians - GIVEN A ROTATION ANGLE IN COMPASS DEGREES, WHICH GO CLOCKWISE,
'                             RETURNS THE EQUIVALENT ROTATION ANGLE IN MATEMATICAL RADIANS, WHICH GO COUNTER-CLOCKWISE
'            ConvertDDtoDMS - GIVEN A DECIMAL DEGREE, RETURNS DEGREES(LONG), MINUTES(LONG) AND SECONDS(DOUBLE)
'            ConvertDMStoDD - GIVEN DEGREES(LONG), MINUTES(LONG) AND SECONDS(DOUBLE), RETURNS DECIMAL DEGREE VALUE(DOUBLE)
' ConvertRotationMathRadiansToCompassDegrees - GIVEN A ROTATION ANGLE IN RADIANS, WHICH GO COUNTER-CLOCKWISE,
'                             RETURNS THE EQUIVALENT ROTATION ANGLE IN COMPASS DEGREES, WHICH GO CLOCKWISE
'ConvertSlopeDegreesToPercent - GIVEN SLOPE VALUE IN DEGREES, RETURNS THE PERCENT SLOPE
'ConvertSlopePercentToDegrees - GIVEN SLOPE IN PERCENT (WHERE 100% = 100), RETURNS THE SLOPE IN DEGREES
'      CreateBoxAroundPoint - GIVEN CENTERPOINT, X-distance, y-distance, RETURNS AN IPolygon RECTANGLE.
'   CreateCircleAroundPoint - GIVEN CENTERPOINT, RADIUS, AND POINT COUNT, RETURNS AN IPolygon CIRCLE.
'  CreateCircleAroundPointGeographic - GIVEN CENTERPOINT, RADIUS, AND POINT COUNT, RETURNS AN IPolygon CIRCLE USING SPHEROIDAL
'                             METHODS
'    CreateCrossAroundPoint - GIVEN CENTERPOINT, HORIZONTAL AND VERTICAL LENGTHS, RETURNS AN IPolyline CROSS
'  CreateEllipseAroundPoint - Given centerpoint, SemiMajor radius, SemiMinor radius, slant angle, and optional number of vertices,
'                             returns an elliptical polygon.
'            CurveToPolygon - SIMILAR TO EllipticArcToPolygon2 EXCEPT THAT IT RETURNS AN IPolygon.  DOESN'T INSERT
'                             POINTS IF SEGMENT IS A LINE
'           CurveToPolyline - SIMILAR TO EllipticArcToPolygon2 EXCEPT THAT IT RETURNS AN IPolyline.  DOESN'T INSERT
'                             POINTS IF SEGMENT IS A LINE
'                  DegToRad - Given Degrees (double), returns Radians (double) (ACCIDENTALLY DUPLICATED THIS FUNCTION!  SEE AsRadians)
'              DegToPercent - Given Slope in Degrees (double), returns Slope in Percent (double) (Note: Percent Slope of 1 = 100%)
'         DistanceHaversine - GIVEN Point A, Point B and optional earth radius, returns distance between points in meters.
'                             Less accurate but faster than using Vincenty's functions.
'  DistanceHaversineNumbers - Identical to DistanceHaversine, but takes Double values for arguments instead of points.  Optionally
'                             also calculates the bearing.
'DistancePythagoreanNumbers - Given coordinates for two points, returns the distances using Pythagorean Theorem
'DistancePointToInfiniteLine - Given 2 consecutive points defining a line with direction, this scripts calculates whether the third point
'                             lies to the right (clockwise) or to the left (counter-clockwise) of the line connecting the first point to
'                             the second point, and the distance from the point to the line.
'    DistancePointToSegment - Given 2 consecutive points defining a segment with direction, this scripts calculates whether the third point
'                             lies to the right (clockwise), left (counter-clockwise) of the line, or on the line connecting the first point to
'                             the second point, and the distance from the point to the segment.  Also optionally gives distance to infinite line
'                             and coordinates of closest point on line.
'    DistanceVincentyPoints - Given 2 points and empty double variables for azimuths, returns distance in meters (using WGS84)
'                             and start and end azimuths of geodesic curve.
'   DistanceVincentyPoints2 - MODIFICATION OF DistanceVincentyPoints, TO ALLOW FOR ANY ELLIPSOID
'   DistanceVincentyNumbers - Given numeric values for latitude and longitude for 2 points, plus empty double variables
'                             for azimuths, returns distance in meters (using WGS84) and start and end azimuths of geodesic curve.
'  DistanceVincentyNumbers2 - MODIFICATION OF DistanceVincentyNumbers, TO ALLOW FOR ANY ELLIPSOID
'      EllipticArcToPolygon - Given a segment collection and number of vertices, returns a polygon4 simulating the ellipse
'                             by generating points along the arc and then calculating a convex hull around the points.
'     EllipticArcToPolygon2 - GIVEN A SEG COLLECTION AND NUMBER OF VERTICES, RETURNS A MULTIPOINT WITH APPROXIMATELY THE
'                             REQUESTED NUMBER OF POINTS DISTRIBUTED ALONG THE ARC.
'  EstimateDistanceOnSphere - GIVEN A GEOGRAPHIC GEOMETRY AND A DISTANCE IN METERS, RETURNS THE NUMBER OF "DEGREES" THAT
'                             NUMBER OF METERS TRANSLATES TO.  GENERATES A NEW POINT THE SPECIFIED DISTANCE FROM THE CENTROID
'                             OF THE GEOMETRY EXTENT, THEN CALCULATES DISTANCE USING PYTHAGOREAN THEOREM.  ONLY USEFUL AS ESTIMATED CONVERSION,
'                             SUCH AS FOR ESTIMATING A PROPER BUFFER DISTANCE
'         EnvelopeToPolygon - GIVEN AN ENVELOPE, RETURNS A POLYGON
'FeaturePlanetOCentricToPlanetOGraphic - RETURNS A NEW GEOMETRY (POINT, POLYLINE, POLYGON OR MULTIPOINT) IN WHICH EACH VERTEX
'                             HAS BEEN SHIFTED FROM AN OCENTRIC TO AN OGRAPHIC LOCATION.
'FeaturePlanetOGraphicToPlanetOCentric - RETURNS A NEW GEOMETRY (POINT, POLYLINE, POLYGON OR MULTIPOINT) IN WHICH EACH VERTEX
'                             HAS BEEN SHIFTED FROM AN OGRAPHIC TO AN OCENTRIC LOCATION.
'ForceAzimuthToCorrectRange - GIVEN A NUMBER, RETURNS A NUMBER IN THE RANGE OF 0 TO <360 DEGREES.  360 IS CONVERTED TO Zero.
'  Graphic_MakeFromGeometry - GIVEN A MAP DOCUMENT, GEOMETRY AND OPTIONAL NAME AND SYMBOLOGY, ADDS GRAPHIC TO MAP.
'Graphic_ReturnElementFromGeometry - GIVEN MAP DOC, GEOMETRY, OPTIONAL NAME AND OPTIONAL ADD-TO-VIEW, RETURNS THE GRAPHIC
'                                    ELEMENT
'                   HArcCos - GIVEN A VALUE, RETURNS INVERSE HYPERBOLIC COSINE
'                   HArcSin - GIVEN A VALUE, RETURNS INVERSE HYPERBOLIC SINE
'                   HArcTan - GIVEN A VALUE, RETURNS INVERSE HYPERBOLIC TANGENT
'                      HCos - GIVEN A VALUE IN RADIANS, RETURNS HYPERBOLIC COSINE
'                      HSin - GIVEN A VALUE IN RADIANS, RETURNS HYPERBOLIC SINE
'                      HTan - GIVEN A VALUE IN RADIANS, RETURNS HYPERBOLIC TANGENT
'                 ModDouble - ACTS LIKE STANDARD MOD FUNCTION, BUT ACCEPTS DOUBLE VALUES AND RETURNS A DOUBLE.
'                             DOES NOT FORCE INPUT VALUES INTO INTEGERS.
'        MultipointCentroid - GIVEN AN IMULTIPOINT, RETURNS AN IPOINT REPRESENTING AVERAGE OF POINTS
'  MultipointCentroidSphere - GIVEN A GEOGRAPHIC IMULTIPOINT, RETURNS AN IPOINT REPRESENTING AVERAGE ON SPHERE
'MultipointCentroidSpheriod - GIVEN A GEOGRAPHIC IMULTIPOINT, RETURNS AN IPOINT REPRESENTING AVERAGE ON SPHEROID
'    MyGeomCheckSpRefDomain - CHECKS A SPATIAL REFERENCE TO SEE WHETHER IT HAS A DOMAIN DEFINED.  RETURNS BOOLEAN
'              PercentToDeg - Given Slope in Percent (double), returns Slope in Degrees (double) (Note: Percent Slope of 1 = 100%)
'                  PointAdd - ADDS TWO POINTS
'         PointLineVincenty - Given Point, Distance and Azimuth, plus empty NewPoint and End Azimuth, plus optional
'                             number of vertices and empty Polyline, returns a new point the specified distance from the
'                             origin along geodesic, plus optional polyline with specified number of vertices.
'        PointLineVincenty2 - MODIFICATION OF PointLineVincenty, TO ALLOW FOR ANY ELLIPSOID
' PointLineVincentyPerPoint - Given pPoint, dblLength, dblAzimuth, empty NewPoint and empty End Azimuth,
'                             fills empty point and azimuth with correct values.
'PointLineVincentyPerPoint2 - MODIFICATION OF PointLineVincentyPerPoint, TO ALLOW FOR ANY ELLIPSOID
'             PointSubtract - SUBTRACTS POINT B FROM POINT A
'         PolygonToPolyline - GIVEN A POLYGON, RETURNS A POLYLINE
'   ProjectToWorldAzimuthal - GIVEN A GEOGRAPHIC POLYGON, RETURNS A POLYGON PROJECTED INTO A CUSTOM WORLD AZIMUTHAL EQUIDSTANT PROJECTION
'                             CENTERED ON POLYGON ENVELOPE CENTROID.
'                  RadToDeg - Given Radians (double), returns Degrees (double) (ACCIDENTALLY DUPLICATED THIS FUNCTION!  SEE AsDegrees)
'    RandomlySelectTriangle - GIVEN A DOUBLE ARRAY LIKE THAT PRODUCED BY TriangulatePolygonToDouble, GENERATES A UNIFORM RANDOM NUMBER
'                             AND FINDS THE CORRECT TRIANGLE INDEX VALUE
'     RandomPointInTriangle - GIVEN 3 PAIRS OF COORDINATES, IT FILLS RandomX AND RandomY VARIABLES WITH COORDINATES THAT EXIST IN
'                             TRIANGLE AND RETURNS A BOOLEAN
'      RandomPointInPolygon - GIVEN AN ARRAY AS PRODUCED BY "TriangulatePolygonToDouble", fills two Double values with Random X and Random Y
'                             coordinates of a point that falls within the polygon.
'     ReturnAngleOfCoverage - GIVEN A POINT AND A POLYLINE OR POLYGON, RETURNS ANGLE OF ARC OF HORIZON OBSCURED BY POLYLINE/POLYGON.
'                             OPTIONALLY RETURNS LEFTMOST AND RIGHTMOST BEARINGS.
'ReturnConvexHullFromFClass - GIVEN AN FLAYER AND OPTIONALLY INSTRUCTIONS TO JUST USE SELECTED FEATURES OR TO APPLY A NEW
'                             ATTRIBUTE QUERY, RETURNS A POLYGON OF THE CONVEX HULL CONTAINING ALL THE FEATURES.
'ReturnConvexHullFromGeometry - GIVEN A GEOMETRY, RETURNS A CONVEX HULL AROUND IT.  CONVEX HULL WILL BE A POLYGON, AND POSSIBLY
'                             AN EMPTY POLYGON IF IT GETS AN EMPTY GEOMETRY.
'ReturnLongestPerpendicularFromSegment - GIVEN COORDINATES OF A SEGMENT AND COORDINATE ARRAY, RETURNS LONGEST DISTANCE
'                             CLOCKWISE AND COUNTERCLOCKWISE FROM THAT SEGMENT.  OPTIONALLY RETURNS COORDINATES OF FARTHEST VERTICES.
'                             THIS FUNCTION EXTENTS SEGMENT TO INFINITE LINE.
'             ReturnMeanDir - GIVEN A DOUBLE ARRAY OF COMPASS DIRECTIONS, RETURNS THE MEAN COMPASS BEARING
'                  ReturnPi - CALCULATES PI USING MACHIN'S FORMULA
'     ReturnWeightedMeanDir - GIVEN A 2-DIMENSIONAL DOUBLE ARRAY OF [COMPASS BEARINGS, WEIGHTS], RETURNS THE MEAN COMPASS BEARING.
'    ReturnWeightedMeanDir2 - GIVEN A 2-DIMENSIONAL DOUBLE ARRAY OF [COMPASS BEARINGS, WEIGHTS], RETURNS THE MEAN COMPASS BEARING PLUS
'                             LOTS OF MEASURES OF DISPERSION.
'       ReturnVonMisesKappa - GIVEN A dblResultantMeanLength (Rho) AND n, RETURNS KAPPA
'ReturnVerticesAsDoubleArray - RETURNS A DOUBLE ARRAY OF X- AND Y-COORDINATES OF ALL VERTICES IN GEOMETRY.  FASTER FOR FUNCTIONS
'                             THAT NEED TO GO THROUGH VERTEX COORDINATES MULTIPLE TIMES.
'SplitMultipartFeatureIntoArray - GIVEN A MULTIPOINT, POLYLINE OR POLYGON, RETURNS AND esriSystem.IArray OF SEPARATE PARTS
'            SolarFunctions - GIVEN A LATITUDE, LONGITUDE, DATE WITH TIME, HOURS DIFFERENT THAN GREENWICH, RETURNS SUNRISE,
'                             SUNSET, SUN DIRECTION AND SUN ANGLE UP AT POINT.
'              ShowVertices - GIVEN A MAP DOC, GEOMETRY AND OPTIONAL NAME, ADDS POINT GRAPHICS TO SCREEN SHOWING WHERE
'                             VERTICES ARE
'    SphericalLatLongToCart - SUBROUTINE:  GIVEN LATITUDE AND LONGITUDE, AND OPTIONAL RADIUS, FILLS X, Y, Z VALUES
'    SphericalCartToLatLong - SUBROUTINE:  GIVEN X, Y, Z, FILLS LATITUDE AND LONGITUDE VALUES
'      SphericalPolygonArea - GIVEN A POLYGON, CALCULATES AREA USING SERIES OF SphericalTriangleArea CALLS
'     SphericalPolygonArea2 - MODIFICATION OF SphericalPolygonArea TO ALLOW USER TO SET CUSTOM ELLIPSOID MAJOR AND MINOR AXES
'   SpheroidalCartToLatLong - GIVEN X, Y, Z, OPTIONAL SPHEROID RADII AND HEIGHT ABOVE ELLPISOID, FILLS LATITUDE AND LONGITUDE VALUES
'   SpheroidalLatLongToCart - GIVEN LATITUDE, LONGITUDE, OPTIONAL SPHEROID RADII AND HEIGHT ABOVE ELLPISOID, FILLS X, Y, Z VALUES
'SpheroidalPolylineFromEndPoints - GIVEN START AND END POINTS IN GEOGRAPHIC COORDINATES, PLUS A NUMBER OF VERTICES, RETURNS
'                             A GEOGRAPHICALLY-PROJECTED POLYLINE WITH THE SPECIFIED NUMBER OF VERTICES EQUALLY SPACED ALONG
'                             THE GREAT CIRCLE ARC CONNECTING THE TWO ENDPOINTS.
'SpheroidalPolylineFromEndPoints2 - REVISION OF SpheroidalPolylineFromEndPoints2 WHICH FIXES A BUG IN WHICH LINES THAT CROSSED
'                             THE DATELINE WOULD HAVE THE LAST POINT ERRONEOUSLY PLACED.
'  SpheroidalPolylineLength - GIVEN A GEOGRAPHIC POLYLINE, RETURNS LENGTH IN METERS BASED ON VINCENTY'S EQUATIONS
' SpheroidalPolylineLength2 - MODIFICATION OF SpheroidalPolylineLength, TO ALLOW FOR ANY ELLIPSOID
'SpheroidalPolylineMidpoint - GIVEN A GEOGRAPHIC POLYLINE, DISTANCE VALUE, booAsRatio, RETURNS POINT AND OPTIONAL POLYLINE DISTANCE
'SpheroidalPolylineMidpoint2 - MODIFICATION OF SpheroidalPolylineMidpoint, TO ALLOW FOR ANY ELLIPSOID
'     SphericalTriangleArea - GIVEN 3 GEOGRAPHIC POINTS, CALCULATES SPHERICAL AREA
'    SphericalTriangleArea2 - MODIFICATION OF SphericalTriangleArea, TO ALLOW USER TO OPTIONALLY SEND CUSTOM MAJOR AND MINOR ELLIPSOID AXES
'   SplitGeometryOnDateLine - GIVEN EITHER A PROJECTED OR GEOGRAPHIC POLYLINE OR POLYGON, WILL SPLIT THE GEOMETRY ON THE DATE-LINE OF THE
'                             GEOGRAPHIC COORDINATE SYSTEM (I.E. THE -180/180 DEGREE LINE).  MIGHT BE USED IN CONJUNCTION WITH
'                             SpheroidalPolylineFromEndPoints2 TO PRODUCE A POLYLINE THAT CORRECTLY INTERSECTS REGIONS AROUND THE
'                             DATELINE.
'SquaredDistanceBetweenSegments - GIVEN 2 MULTIDIMENSIAL ARRAYS FOR THE START AND END OF A SEGMENT, AND 2 MORE FOR ANOTHER SEGMENT, THIS
'                             FUNCTION WILL FILL 2 X-DIMENSIONAL ARRAYS WITH THE CLOSEST POINT COORDINATES ON EACH SEGMENT PLUS RETURN
'                             THE SQUARED DISTANCE BETWEEN THE TWO SEGMENTS
'        TriangleAreaPoints - GIVEN THREE POINTS, RETURNS AREA IN LOCAL UNITS
'        TriangleAreaPoints - GIVEN PAIRS OF POINT X/Y COORDINATES, RETURNS AREA IN LOCAL UNITS
'      TriangleAreaPoints3D - GIVEN THREE 3-DIMENSIONAL POINTS, RETURNS AREA IN LOCAL UNITS
'        TriangleCentroid3D - GIVEN X,Y,Z COORDINATES FOR THREE 3D POINTS, RETURNS 3D CENTROID
'     TriangleCentroidPlane - GIVEN X,Y COORDINATES FOR THREE POINTS, RETURNS CENTROID
'TriangulatePolygonToDouble - GIVEN A POLYGON, RETURNS A DOUBLE ARRAY WITH 6x[Triangle Count] DIMENSIONS, WITH 1ST COLUMN
'                             HOLDING CUMULATIVE PROPORTIONAL AREA AND THE OTHER 6 COLUMNS CONTAINING VERTEX X/Y COORDINATES
'           UnionGeometries - GIVEN A VARIANT ARRAY OF GEOMETRY OBJECTS, RETURNS A SINGLE UNIONED VERSION WITH THE SAME DIMENSION
'          UnionGeometries2 - GIVEN A VARIANT ARRAY OF GEOMETRY OBJECTS, RETURNS A SINGLE UNIONED VERSION WITH THE SAME DIMENSION.
'                             INCLUDES THE OPTION TO SET A MAXIMUM NUMBER OF GEOMETRIES TO UNION.
'      XYOCentricToOGraphic - GIVEN LONGITUDE, LATITUDE, ELLIPSOID MAJOR AND MINOR AXES, AND OPTIONAL LONGITUDE SHIFT, SETS NEW
'                             LATITUDE AND LONGITUDE VALUES BY CONVERTING FROM OCENTRIC TO OGRAPHIC.
'      XYOGraphicToOCentric - GIVEN LONGITUDE, LATITUDE, ELLIPSOID MAJOR AND MINOR AXES, AND OPTIONAL LONGITUDE SHIFT, SETS NEW
'                             LATITUDE AND LONGITUDE VALUES BY CONVERTING FROM OGRAPHIC TO OCENTRIC.


Option Explicit
   
Public Enum JenSphericalMethod
  ENUM_UseTrigonometry = 1
  ENUM_UseSpherical = 2
  ENUM_UseSpheroidal = 4
End Enum


Public Enum JenClockwiseConstants
  ENUM_CounterClockwise = 0
  Enum_OnLine = 1
  Enum_Clockwise = 2
End Enum

Public Enum JenSolarConditions
  ENUM_SunriseAndSunset = 1
  ENUM_AlwaysNight = 2
  ENUM_AlwaysDay = 4
End Enum

Const dblPI As Double = 3.14159265358979
Const dblE As Double = 2.71828182845905

Public Sub SolarFunctions(dblLatitude As Double, dblLongitude As Double, datDateWithTime As Date, _
    dblHoursFromGreenwich As Double, Optional lngSunriseExists As JenSolarConditions, _
    Optional dblSunrise As Double, Optional dblSunset As Double, _
    Optional dblSunDirection As Double, Optional dblSunAngleUp As Double, _
    Optional dblSunDirectionAtSunrise As Double = -9999, Optional dblSunDirectionAtSunset As Double = -9999)
  
  ' ADAPTED FROM http://www.esrl.noaa.gov/gmd/grad/solcalc/
  ' SAMPLE EXCEL FILE http://www.esrl.noaa.gov/gmd/grad/solcalc/NOAA_Solar_Calculations_day.xls
  ' GLOSSARY OF TERMS AT http://www.esrl.noaa.gov/gmd/grad/solcalc/glossary.html
  ' Sample Code at Bottom
  
  ' VARIABLES BELOW ARE NAMED ACCORDING TO DESCRIPTION AND EXCEL COLUMN
  
  Dim dblA As Double
  Dim dblB As Double
  Dim boo_W_Crashed As Boolean
  Dim boo_Y_Crashed As Boolean
  Dim boo_Z_Crashed As Boolean
  Dim boo_AA_Crashed As Boolean
  
  Dim dbl_E_Time_PastLocalMidnight As Double
  Dim dbl_F_JulianDay As Double
  Dim dbl_G_Julian_Century As Double
  Dim dbl_I_Geom_Mean_Long_Sun_Deg As Double
  Dim dbl_J_GeomMean_Anom_Sun_Deg As Double
  Dim dbl_K_Eccent_Earth_Orbit As Double
  Dim dbl_L_Sun_Eq_of_Ctr As Double
  Dim dbl_M_Sun_True_Long_Deg As Double
  Dim dbl_N_Sun_True_Anom_Deg As Double
  Dim dbl_O_Sun_Rad_vector_AUs As Double
  Dim dbl_P_Sun_App_Long_Deg As Double
  Dim dbl_Q_Mean_Obliq_Ecliptic_Deg As Double
  Dim dbl_R_Obliq_Corr_Deg As Double
  Dim dbl_S_Sun_Rt_Ascen_Deg As Double
  Dim dbl_T_Sun_Declin_Deg As Double
  Dim dbl_U_Var_Y As Double
  Dim dbl_V_EqOfTime_Minutes As Double
  Dim dbl_W_AH_Sunrise_Deg As Double
  Dim dbl_X_Solar_Noon_LST As Double
  Dim dbl_Y_Sunrise_Time_LST As Double
  Dim dbl_Z_Sunset_Time_LST As Double
  Dim dbl_AA_Sunlight_Duration_Min As Double
  Dim dbl_AB_True_Solar_Time_Min As Double
  Dim dbl_AC_Hour_Angle_Deg As Double
  Dim dbl_AD_Solar_Zenith_Angle_Deg As Double
  Dim dbl_AE_Solar_Elevation_Angle_Deg As Double
  Dim dbl_AF_Approx_Atmospheric_Refraction_Deg As Double
  Dim dbl_AG_Solar_Elev_Corrected_for_Refract_Deg As Double
  Dim dbl_AH_Solar_Azimuth_Angle_Deg_CW_From_N As Double
  
  ' ALL REFERENCE EQUATIONS BELOW ARE COPIED DIRECTLY FROM EXCEL.
  ' ALL REFERENCE VARIABLES HAVE "2" IN THE NAME BECAUSE THEY WERE COPIED FROM ROW 2.
  ' BE CAREFUL OF EXCEL "ATAN2" FUNCTION BECAUSE IT USES NON-TRADITIONAL PARAMETER ORDER.
  ' BE CAREFUL OF EXCEL "MOD" FUNCTION BECAUSE IT RETURNS DOUBLE VALUES, NOT INTEGER VALUES LIKE VB MOD.
  
  ' $B$3 = Latitude
  ' $B$4 = Longitude
  ' $B$5 = hours difference from Greenwich
  ' $B$7 = Date
  
  ' SOME FUNCTIONS FILL FAIL IF NO SUNRISE OR SUNSET ON A PARTICULAR DAY.
  '  dbl_W_AH_Sunrise_Deg
  '  dbl_Y_Sunrise_Time_LST
  '  dbl_Z_Sunset_Time_LST
  '  dbl_AA_Sunlight_Duration_Min
  ' SHOULD BE ABLE TO CATCH THESE AND SAY WHETHER IT IS CONSTANT DAYLIGHT OR NIGHT BASED ON
  '   SOLAR ELEVATION AT SOLAR NOON.  NEGATIVE VALUE MEANS NIGHT.
  
  'E2 = 0.1/24, E3 = E2+0.1/24, E4 = E3+0.1/24, etc. to increase in 6-minute increments
  ' BASICALLY THE NUMBER OF DAYS PAST MIDNIGHT, SO WILL ALWAYS BE < 1.
  dbl_E_Time_PastLocalMidnight = CDbl(datDateWithTime) - Fix(datDateWithTime)
'   Debug.Print "dbl_E_Time_PastLocalMidnight = " & Format(dbl_E_Time_PastLocalMidnight, "0.000000000000")
  
  'F2 = D2+2415018.5+E2-$B$5/24
  dbl_F_JulianDay = CDbl(datDateWithTime) + 2415018.5 - (dblHoursFromGreenwich / 24)
'   Debug.Print "dbl_F_JulianDay = " & CStr(dbl_F_JulianDay)
  
  'G2 =(F2-2451545)/36525
  dbl_G_Julian_Century = (dbl_F_JulianDay - 2451545) / 36525
'   Debug.Print "dbl_G_Julian_Century = " & CStr(dbl_G_Julian_Century)
  
  'I2 =MOD(280.46646+G2*(36000.76983 + G2*0.0003032),360)
  dbl_I_Geom_Mean_Long_Sun_Deg = _
      ModDouble(280.46646 + dbl_G_Julian_Century * (36000.76983 + dbl_G_Julian_Century * 0.0003032), 360)
'   Debug.Print "dbl_I_Geom_Mean_Long_Sun_Deg = " & Format(dbl_I_Geom_Mean_Long_Sun_Deg, "0.000000000000")
  
  'J2 =357.52911+G2*(35999.05029 - 0.0001537*G2)
  dbl_J_GeomMean_Anom_Sun_Deg = 357.52911 + dbl_G_Julian_Century * (35999.05029 - 0.0001537 * dbl_G_Julian_Century)
'   Debug.Print "dbl_J_GeomMean_Anom_Sun_Deg = " & Format(dbl_J_GeomMean_Anom_Sun_Deg, "0.000000000000")
  
  'K2 =0.016708634-G2*(0.000042037+0.0000001267*G2)
  dbl_K_Eccent_Earth_Orbit = 0.016708634 - dbl_G_Julian_Century * (0.000042037 + 0.0000001267 * dbl_G_Julian_Century)
'   Debug.Print "dbl_K_Eccent_Earth_Orbit = " & Format(dbl_K_Eccent_Earth_Orbit, "0.000000000000")
  
  'L2 =SIN(RADIANS(J2))*(1.914602-G2*(0.004817+0.000014*G2))+SIN(RADIANS(2*J2))*(0.019993-0.000101*G2)+SIN(RADIANS(3*J2))*0.000289
  dbl_L_Sun_Eq_of_Ctr = Sin(AsRadians(dbl_J_GeomMean_Anom_Sun_Deg)) * (1.914602 - dbl_G_Julian_Century * _
      (0.004817 + 0.000014 * dbl_G_Julian_Century)) + Sin(AsRadians(2 * dbl_J_GeomMean_Anom_Sun_Deg)) * _
      (0.019993 - 0.000101 * dbl_G_Julian_Century) + Sin(AsRadians(3 * dbl_J_GeomMean_Anom_Sun_Deg)) * 0.000289
'   Debug.Print "dbl_L_Sun_Eq_of_Ctr = " & Format(dbl_L_Sun_Eq_of_Ctr, "0.000000000000")
  
  'M2 =I2+L2
  dbl_M_Sun_True_Long_Deg = dbl_I_Geom_Mean_Long_Sun_Deg + dbl_L_Sun_Eq_of_Ctr
'   Debug.Print "dbl_M_Sun_True_Long_Deg = " & Format(dbl_M_Sun_True_Long_Deg, "0.000000000000")
  
  'N2 =J2+L2
  dbl_N_Sun_True_Anom_Deg = dbl_J_GeomMean_Anom_Sun_Deg + dbl_L_Sun_Eq_of_Ctr
'   Debug.Print "dbl_N_Sun_True_Anom_Deg = " & Format(dbl_N_Sun_True_Anom_Deg, "0.000000000000")
  
  'O2 =(1.000001018*(1-K2*K2))/(1+K2*COS(RADIANS(N2)))
  dbl_O_Sun_Rad_vector_AUs = (1.000001018 * (1 - dbl_K_Eccent_Earth_Orbit * dbl_K_Eccent_Earth_Orbit)) / _
      (1 + dbl_K_Eccent_Earth_Orbit * Cos(AsRadians(dbl_N_Sun_True_Anom_Deg)))
'   Debug.Print "dbl_O_Sun_Rad_vector_AUs = " & Format(dbl_O_Sun_Rad_vector_AUs, "0.000000000000")
  
  'P2 =M2-0.00569-0.00478*SIN(RADIANS(125.04-1934.136*G2))
  dbl_P_Sun_App_Long_Deg = dbl_M_Sun_True_Long_Deg - 0.00569 - 0.00478 * _
      Sin(AsRadians(125.04 - 1934.136 * dbl_G_Julian_Century))
'   Debug.Print "dbl_P_Sun_App_Long_Deg = " & Format(dbl_P_Sun_App_Long_Deg, "0.000000000000")
  
  'Q2 =23+(26+((21.448-G2*(46.815+G2*(0.00059-G2*0.001813))))/60)/60
  dbl_Q_Mean_Obliq_Ecliptic_Deg = 0.00059 - (dbl_G_Julian_Century * 0.001813)
  dbl_Q_Mean_Obliq_Ecliptic_Deg = 46.815 + (dbl_G_Julian_Century * dbl_Q_Mean_Obliq_Ecliptic_Deg)
  dbl_Q_Mean_Obliq_Ecliptic_Deg = 21.448 - (dbl_G_Julian_Century * dbl_Q_Mean_Obliq_Ecliptic_Deg)
  dbl_Q_Mean_Obliq_Ecliptic_Deg = 23 + ((26 + (dbl_Q_Mean_Obliq_Ecliptic_Deg / 60)) / 60)
'   Debug.Print "dbl_Q_Mean_Obliq_Ecliptic_Deg = " & CStr(dbl_Q_Mean_Obliq_Ecliptic_Deg)
  
  'R2 =Q2+0.00256*COS(RADIANS(125.04-1934.136*G2))
  dbl_R_Obliq_Corr_Deg = dbl_Q_Mean_Obliq_Ecliptic_Deg + 0.00256 * _
      Cos(AsRadians(125.04 - 1934.136 * dbl_G_Julian_Century))
'   Debug.Print "dbl_R_Obliq_Corr_Deg = " & Format(dbl_R_Obliq_Corr_Deg, "0.000000000000")
  
  'S2 =DEGREES(ATAN2(COS(RADIANS(P2)),COS(RADIANS(R2))*SIN(RADIANS(P2))))
  ' NOTE:  EXCEL USES UNUSUAL ATAN2 DEFINITION.  I SWITCHED PARAMETERS IN MY FUNCTION
  dbl_S_Sun_Rt_Ascen_Deg = AsDegrees(atan2 _
      (Cos(AsRadians(dbl_R_Obliq_Corr_Deg)) * Sin(AsRadians(dbl_P_Sun_App_Long_Deg)), _
      Cos(AsRadians(dbl_P_Sun_App_Long_Deg))))
'   Debug.Print "dbl_S_Sun_Rt_Ascen_Deg = " & Format(dbl_S_Sun_Rt_Ascen_Deg, "0.000000000000")
  
  'T2 =DEGREES(ASIN(SIN(RADIANS(R2))*SIN(RADIANS(P2))))
  dbl_T_Sun_Declin_Deg = AsDegrees(ArcSinJen(Sin(AsRadians(dbl_R_Obliq_Corr_Deg)) * _
      Sin(AsRadians(dbl_P_Sun_App_Long_Deg))))
'   Debug.Print "dbl_T_Sun_Declin_Deg = " & Format(dbl_T_Sun_Declin_Deg, "0.000000000000")
  
  'U2 =TAN(RADIANS(R2/2))*TAN(RADIANS(R2/2))
  dbl_U_Var_Y = Tan(AsRadians(dbl_R_Obliq_Corr_Deg / 2)) * Tan(AsRadians(dbl_R_Obliq_Corr_Deg / 2))
'   Debug.Print "dbl_U_Var_Y = " & Format(dbl_U_Var_Y, "0.000000000000")
    
  'V2 =4*DEGREES(U2*SIN(2*RADIANS(I2))-2*K2*SIN(RADIANS(J2))+4*K2*U2*SIN(RADIANS(J2))*COS(2*RADIANS(I2))-0.5*U2*U2*SIN(4*RADIANS(I2))-1.25*K2*K2*SIN(2*RADIANS(J2)))
  dbl_V_EqOfTime_Minutes = dbl_U_Var_Y * Sin(2 * AsRadians(dbl_I_Geom_Mean_Long_Sun_Deg))
  dbl_V_EqOfTime_Minutes = dbl_V_EqOfTime_Minutes - _
      (2 * dbl_K_Eccent_Earth_Orbit * Sin(AsRadians(dbl_J_GeomMean_Anom_Sun_Deg)))
  dbl_V_EqOfTime_Minutes = dbl_V_EqOfTime_Minutes + _
      4 * dbl_K_Eccent_Earth_Orbit * dbl_U_Var_Y * Sin(AsRadians(dbl_J_GeomMean_Anom_Sun_Deg)) * _
      Cos(2 * AsRadians(dbl_I_Geom_Mean_Long_Sun_Deg))
  dbl_V_EqOfTime_Minutes = dbl_V_EqOfTime_Minutes - _
      0.5 * dbl_U_Var_Y * dbl_U_Var_Y * Sin(4 * AsRadians(dbl_I_Geom_Mean_Long_Sun_Deg))
  dbl_V_EqOfTime_Minutes = dbl_V_EqOfTime_Minutes - _
      1.25 * dbl_K_Eccent_Earth_Orbit * dbl_K_Eccent_Earth_Orbit * Sin(2 * AsRadians(dbl_J_GeomMean_Anom_Sun_Deg))
  dbl_V_EqOfTime_Minutes = 4 * AsDegrees(dbl_V_EqOfTime_Minutes)
'   Debug.Print "dbl_V_EqOfTime_Minutes = " & Format(dbl_V_EqOfTime_Minutes, "0.000000000000")

  'W2 =DEGREES(ACOS(COS(RADIANS(90.833))/(COS(RADIANS($B$3))*COS(RADIANS(T2)))-TAN(RADIANS($B$3))*TAN(RADIANS(T2))))
  ' NOTE:  THIS VALUE COULD CRASH IF NO SUNRISE OR SUNSET; PAST ARCTIC OR ANTARCTIC CIRCLE AND AT THE RIGHT TIME OF YEAR
'  dbl_W_AH_Sunrise_Deg = Cos(AsRadians(90.833)) / _
'      (Cos(AsRadians(dblLatitude)) * Cos(AsRadians(dbl_T_Sun_Declin_Deg)))
'''   Debug.Print "dbl_W_AH_Sunrise_Deg: A = " & Format(dbl_W_AH_Sunrise_Deg, "0.000000000000")
'  dbl_W_AH_Sunrise_Deg = dbl_W_AH_Sunrise_Deg - (Tan(AsRadians(dblLatitude)) * Tan(AsRadians(dbl_T_Sun_Declin_Deg)))
'''   Debug.Print "dbl_W_AH_Sunrise_Deg: B = " & Format(dbl_W_AH_Sunrise_Deg, "0.000000000000")
'  dbl_W_AH_Sunrise_Deg = AsDegrees(ArcCosJen(dbl_W_AH_Sunrise_Deg))
''  dbl_W_AH_Sunrise_Deg = AsDegrees(ArcCosJen(Cos(AsRadians(90.833)) / _
''      (Cos(AsRadians(dblLatitude)) * Cos(AsRadians(dbl_T_Sun_Declin_Deg))) - _
'      Tan(AsRadians(dblLatitude)) * Tan(AsRadians(dbl_T_Sun_Declin_Deg))))
  
  dbl_W_AH_Sunrise_Deg = Return_W_AH_Sunrise_Deg(dblLatitude, dbl_T_Sun_Declin_Deg, boo_W_Crashed)
'  Debug.Print "dbl_W_AH_Sunrise_Deg = " & Format(dbl_W_AH_Sunrise_Deg, "0.000000000000")

  'X2 =(720-4*$B$4-V2+$B$5*60)/1440
  dbl_X_Solar_Noon_LST = (720 - 4 * dblLongitude - dbl_V_EqOfTime_Minutes + dblHoursFromGreenwich * 60) / 1440
'   Debug.Print "dbl_X_Solar_Noon_LST = " & Format(dbl_X_Solar_Noon_LST, "Hh:Nn:Ss")
  
  If boo_W_Crashed Then  ' Sunrise, Sunset and Sun Duration will also crash
    
    dbl_Y_Sunrise_Time_LST = -9999
    dbl_Z_Sunset_Time_LST = -9999
    dbl_AA_Sunlight_Duration_Min = -9999
    
  Else
    'Y2 =X2-W2*4/1440
    dbl_Y_Sunrise_Time_LST = dbl_X_Solar_Noon_LST - dbl_W_AH_Sunrise_Deg * 4 / 1440
  '   Debug.Print "dbl_Y_Sunrise_Time_LST = " & Format(dbl_Y_Sunrise_Time_LST, "Hh:Nn:Ss")
  
    'Z2 =X2+W2*4/1440
    dbl_Z_Sunset_Time_LST = dbl_X_Solar_Noon_LST + dbl_W_AH_Sunrise_Deg * 4 / 1440
  '   Debug.Print "dbl_Z_Sunset_Time_LST = " & Format(dbl_Z_Sunset_Time_LST, "Hh:Nn:Ss")
  
    'AA2 =8*W2
    dbl_AA_Sunlight_Duration_Min = 8 * dbl_W_AH_Sunrise_Deg
  '   Debug.Print "dbl_AA_Sunlight_Duration_Min = " & Format(dbl_AA_Sunlight_Duration_Min, "0.000000000000")
  End If
  
  'AB2 =MOD(E2*1440+V2+4*$B$4-60*$B$5,1440)
  dbl_AB_True_Solar_Time_Min = ModDouble(dbl_E_Time_PastLocalMidnight * 1440 + dbl_V_EqOfTime_Minutes + _
      4 * dblLongitude - 60 * dblHoursFromGreenwich, 1440)
'   Debug.Print "dbl_AB_True_Solar_Time_Min = " & Format(dbl_AB_True_Solar_Time_Min, "0.000000000000")

  'AC2 =IF(AB2/4<0,AB2/4+180,AB2/4-180)
  If dbl_AB_True_Solar_Time_Min / 4 < 0 Then
    dbl_AC_Hour_Angle_Deg = dbl_AB_True_Solar_Time_Min / 4 + 180
  Else
    dbl_AC_Hour_Angle_Deg = dbl_AB_True_Solar_Time_Min / 4 - 180
  End If
'   Debug.Print "dbl_AC_Hour_Angle_Deg = " & Format(dbl_AC_Hour_Angle_Deg, "0.000000000000")

  'AD2 =DEGREES(ACOS(SIN(RADIANS($B$3))*SIN(RADIANS(T2))+COS(RADIANS($B$3))*COS(RADIANS(T2))*COS(RADIANS(AC2))))
  ' ZENITH ANGLE IS MEASURED DOWN FROM STRAIGHT UP
  dbl_AD_Solar_Zenith_Angle_Deg = AsDegrees(ArcCosJen(Sin(AsRadians(dblLatitude)) * Sin(AsRadians(dbl_T_Sun_Declin_Deg)) + _
      Cos(AsRadians(dblLatitude)) * Cos(AsRadians(dbl_T_Sun_Declin_Deg)) * Cos(AsRadians(dbl_AC_Hour_Angle_Deg))))
'   Debug.Print "dbl_AD_Solar_Zenith_Angle_Deg = " & Format(dbl_AD_Solar_Zenith_Angle_Deg, "0.000000000000")

  'AE2 =90-AD2
  ' THIS IS THE TRUE SOLAR ELEVATION; REGARDLESS OF WHERE WE SEE IT
  dbl_AE_Solar_Elevation_Angle_Deg = 90 - dbl_AD_Solar_Zenith_Angle_Deg
'   Debug.Print "dbl_AE_Solar_Elevation_Angle_Deg = " & Format(dbl_AE_Solar_Elevation_Angle_Deg, "0.000000000000")

  'AF2 =IF(AE2>85,0,IF(AE2>5,58.1/TAN(RADIANS(AE2))-0.07/POWER(TAN(RADIANS(AE2)),3)+0.000086/POWER(TAN(RADIANS(AE2)),5),IF(AE2>-0.575,1735+AE2*(-518.2+AE2*(103.4+AE2*(-12.79+AE2*0.711))),-20.772/TAN(RADIANS(AE2)))))/3600
  If dbl_AE_Solar_Elevation_Angle_Deg > 85 Then
    dbl_AF_Approx_Atmospheric_Refraction_Deg = 0
  Else
    If dbl_AE_Solar_Elevation_Angle_Deg > 5 Then
      ' IF(AE2>5,58.1/TAN(RADIANS(AE2))-0.07/POWER(TAN(RADIANS(AE2)),3)+0.000086/POWER(TAN(RADIANS(AE2)),5)
      dbl_AF_Approx_Atmospheric_Refraction_Deg = 58.1 / Tan(AsRadians(dbl_AE_Solar_Elevation_Angle_Deg)) - 0.07 / _
           (Tan(AsRadians(dbl_AE_Solar_Elevation_Angle_Deg))) ^ 3 + _
           0.000086 / ((Tan(AsRadians(dbl_AE_Solar_Elevation_Angle_Deg))) ^ 5)
    Else
      ' IF(AE2>-0.575,1735+AE2*(-518.2+AE2*(103.4+AE2*(-12.79+AE2*0.711))),-20.772/TAN(RADIANS(AE2)))))/3600
      
      If dbl_AE_Solar_Elevation_Angle_Deg > -0.575 Then
        dbl_AF_Approx_Atmospheric_Refraction_Deg = (-518.2 + dbl_AE_Solar_Elevation_Angle_Deg * _
            (103.4 + dbl_AE_Solar_Elevation_Angle_Deg * (-12.79 + dbl_AE_Solar_Elevation_Angle_Deg * 0.711)))
        dbl_AF_Approx_Atmospheric_Refraction_Deg = 1735 + dbl_AE_Solar_Elevation_Angle_Deg * _
            dbl_AF_Approx_Atmospheric_Refraction_Deg

      Else
        dbl_AF_Approx_Atmospheric_Refraction_Deg = -20.772 / Tan(AsRadians(dbl_AE_Solar_Elevation_Angle_Deg))
      End If
    End If
  End If
  dbl_AF_Approx_Atmospheric_Refraction_Deg = dbl_AF_Approx_Atmospheric_Refraction_Deg / 3600
'   Debug.Print "dbl_AF_Approx_Atmospheric_Refraction_Deg = " & Format(dbl_AF_Approx_Atmospheric_Refraction_Deg, "0.000000000000")

  'AG2 =AE2+AF2
  ' THIS IS WHERE WE SEE THE SUN; WE SEE IT BEFORE IT HAS ACTUALLY COME UP OVER THE HORIZON.
  dbl_AG_Solar_Elev_Corrected_for_Refract_Deg = _
      dbl_AE_Solar_Elevation_Angle_Deg dbl_AF_Approx_Atmospheric_Refraction_Deg
'   Debug.Print "dbl_AG_Solar_Elev_Corrected_for_Refract_Deg = " & Format(dbl_AG_Solar_Elev_Corrected_for_Refract_Deg, "0.000000000000")

'  'AH2 =IF(AC2>0,MOD(DEGREES(ACOS(((SIN(RADIANS($B$3))*COS(RADIANS(AD2)))-SIN(RADIANS(T2)))/(COS(RADIANS($B$3))*SIN(RADIANS(AD2)))))+180,360),MOD(540-DEGREES(ACOS(((SIN(RADIANS($B$3))*COS(RADIANS(AD2)))-SIN(RADIANS(T2)))/(COS(RADIANS($B$3))*SIN(RADIANS(AD2))))),360))
  If dbl_AC_Hour_Angle_Deg > 0 Then
    ' MOD(DEGREES(ACOS(((SIN(RADIANS($B$3))*COS(RADIANS(AD2)))-SIN(RADIANS(T2)))/(COS(RADIANS($B$3))*SIN(RADIANS(AD2)))))+180,360)
    dblA = Sin(AsRadians(dblLatitude)) * Cos(AsRadians(dbl_AD_Solar_Zenith_Angle_Deg)) - _
        Sin(AsRadians(dbl_T_Sun_Declin_Deg))
    dblB = Cos(AsRadians(dblLatitude)) * Sin(AsRadians(dbl_AD_Solar_Zenith_Angle_Deg))
    dbl_AH_Solar_Azimuth_Angle_Deg_CW_From_N = ModDouble(AsDegrees(ArcCosJen(dblA / dblB)) + 180, 360)
  Else
    ' MOD(540-DEGREES(ACOS(((SIN(RADIANS($B$3))*COS(RADIANS(AD2)))-SIN(RADIANS(T2)))/(COS(RADIANS($B$3))*SIN(RADIANS(AD2))))),360))
    dblA = (Sin(AsRadians(dblLatitude)) * Cos(AsRadians(dbl_AD_Solar_Zenith_Angle_Deg))) - _
        Sin(AsRadians(dbl_T_Sun_Declin_Deg))
    dblB = Cos(AsRadians(dblLatitude)) * Sin(AsRadians(dbl_AD_Solar_Zenith_Angle_Deg))
    dbl_AH_Solar_Azimuth_Angle_Deg_CW_From_N = ModDouble( _
        540 - AsDegrees(ArcCosJen(dblA / dblB)), 360)
  End If
'   Debug.Print "dbl_AH_Solar_Azimuth_Angle_Deg_CW_From_N = " & Format(dbl_AH_Solar_Azimuth_Angle_Deg_CW_From_N, "0.000000000000")
  
  If boo_W_Crashed Then
    If dbl_AG_Solar_Elev_Corrected_for_Refract_Deg > 0 Then
      lngSunriseExists = ENUM_AlwaysDay
    Else
      lngSunriseExists = ENUM_AlwaysNight
    End If
  Else
    lngSunriseExists = ENUM_SunriseAndSunset
  End If
  
  dblSunrise = dbl_Y_Sunrise_Time_LST
  dblSunset = dbl_Z_Sunset_Time_LST
  dblSunDirection = dbl_AH_Solar_Azimuth_Angle_Deg_CW_From_N
  dblSunAngleUp = dbl_AG_Solar_Elev_Corrected_for_Refract_Deg
  
  
  Dim datFullSunriseDate As Date
  Dim datFullSunsetDate As Date
  
  If dblSunDirectionAtSunrise <> -9999 Then
    If boo_W_Crashed Then
      dblSunDirectionAtSunrise = -9999
    Else
      If CDbl(datDateWithTime) < 0 Then
        datFullSunriseDate = CDate(CDbl(Fix(datDateWithTime)) - dblSunrise)
      Else
        datFullSunriseDate = CDate(CDbl(Fix(datDateWithTime)) + dblSunrise)
      End If
      SolarFunctions dblLatitude, dblLongitude, datFullSunriseDate, dblHoursFromGreenwich, , _
          , , dblSunDirectionAtSunrise
    End If
  End If
    
  If dblSunDirectionAtSunset <> -9999 Then
    If boo_W_Crashed Then
      dblSunDirectionAtSunset = -9999
    Else
      If CDbl(datDateWithTime) < 0 Then
        datFullSunriseDate = CDate(CDbl(Fix(datDateWithTime)) - dblSunset)
      Else
        datFullSunriseDate = CDate(CDbl(Fix(datDateWithTime)) + dblSunset)
      End If
      SolarFunctions dblLatitude, dblLongitude, datFullSunriseDate, dblHoursFromGreenwich, , _
          , , dblSunDirectionAtSunset
    End If
  End If
  
  ' SAMPLE CODE
  '  Debug.Print "--------------------------------------"
  '
  '  Dim dblLatitude As Double
  '  Dim dblLongitude As Double
  '  Dim datDateWithTime As Date
  '  Dim dblHoursFromGreenwich As Double
  '
  '
  '  Dim dblSunrise As Double
  '  Dim dblSunset As Double
  '  Dim dblSunDirection As Double
  '  Dim dblSunAngleUp As Double
  '  Dim lngSolarOption As JenSolarConditions
  '  Dim dblTimePastMidnight As Double
  '  Dim dblSunDirectionAtSunrise As Double
  '  Dim dblSunDirectionAtSunset As Double
  '
  '  dblLatitude = 34.98
  '  dblLongitude = -111.60592
  '  datDateWithTime = CDate("6/21/2010 20:00:00")
  ''  datDateWithTime = DateAdd("h", 7, Now)
  '  dblHoursFromGreenwich = -7
  ''  dblHoursFromGreenwich = 0
  '
  '  Debug.Print "Date as Double = " & CDbl(datDateWithTime)
  '  Debug.Print Format(datDateWithTime, "Long Date")
  '  Debug.Print Format(datDateWithTime, "Long Time")
  '  Debug.Print "Longitude = " & CStr(dblLongitude)
  '  Debug.Print "Latitude = " & CStr(dblLatitude)
  '
  '  SolarFunctions dblLatitude, dblLongitude, datDateWithTime, dblHoursFromGreenwich, _
  '     lngSolarOption, dblSunrise, dblSunset, dblSunDirection, dblSunAngleUp, dblSunDirectionAtSunrise, _
  '     dblSunDirectionAtSunset
  '
  '  dblTimePastMidnight = CDbl(datDateWithTime) - Fix(datDateWithTime)
  '
  '  Debug.Print "---"
  '  If lngSolarOption = ENUM_SunriseAndSunset Then
  '    Debug.Print "Sunrise = " & Format(dblSunrise, "Hh:Nn:Ss")
  '    Debug.Print "Sunset = " & Format(dblSunset, "Hh:Nn:Ss")
  '    Debug.Print "Observed Time >= Sunrise = " & Format(dblTimePastMidnight >= dblSunrise, ">")
  '    Debug.Print "Observed Time <= Sunset = " & Format(dblTimePastMidnight <= dblSunset, ">")
  '    Debug.Print "Sun Visible at Time = " & Format((dblTimePastMidnight >= dblSunrise) And _
  '          (dblTimePastMidnight <= dblSunset), ">")
  '  Else
  '    If lngSolarOption = ENUM_AlwaysDay Then
  '      Debug.Print "No sunrise or sunset; Always day..."
  '    Else
  '      Debug.Print "No sunrise or sunset; Always Night..."
  '    End If
  '  End If
  '
  '  Debug.Print "Sun Direction = " & CStr(dblSunDirection) & " degrees"
  '  Debug.Print "Sun Angle = " & CStr(dblSunAngleUp) & " degrees up"
  '  Debug.Print "Sun Direction at Sunrise = " & CStr(dblSunDirectionAtSunrise) & " degrees"
  '  Debug.Print "Sun Direction at Sunset = " & CStr(dblSunDirectionAtSunset) & " degrees"
  '  Debug.Print "---"
  '  Debug.Print "Done..."
  
  dblSunrise = dbl_Y_Sunrise_Time_LST
  dblSunset = dbl_Z_Sunset_Time_LST
  dblSunDirection = dbl_AH_Solar_Azimuth_Angle_Deg_CW_From_N
  dblSunAngleUp = dbl_AG_Solar_Elev_Corrected_for_Refract_Deg
  
End Sub

Private Function Return_W_AH_Sunrise_Deg(dblLatitude As Double, dbl_T_Sun_Declin_Deg As Double, _
     booCrashed As Boolean) As Double
  On Error GoTo FunctionFailed
  
  Dim dbl_W_AH_Sunrise_Deg As Double
  
  booCrashed = False
  'W2 =DEGREES(ACOS(COS(RADIANS(90.833))/(COS(RADIANS($B$3))*COS(RADIANS(T2)))-TAN(RADIANS($B$3))*TAN(RADIANS(T2))))
  ' NOTE:  THIS VALUE COULD CRASH IF NO SUNRISE OR SUNSET; PAST ARCTIC OR ANTARCTIC CIRCLE AND AT THE RIGHT TIME OF YEAR
  dbl_W_AH_Sunrise_Deg = Cos(AsRadians(90.833)) / _
      (Cos(AsRadians(dblLatitude)) * Cos(AsRadians(dbl_T_Sun_Declin_Deg)))
  dbl_W_AH_Sunrise_Deg = dbl_W_AH_Sunrise_Deg - (Tan(AsRadians(dblLatitude)) * Tan(AsRadians(dbl_T_Sun_Declin_Deg)))
  dbl_W_AH_Sunrise_Deg = AsDegrees(ArcCosJen(dbl_W_AH_Sunrise_Deg))
  Return_W_AH_Sunrise_Deg = dbl_W_AH_Sunrise_Deg
  
  Exit Function
  
FunctionFailed:
  booCrashed = True
  Return_W_AH_Sunrise_Deg = -9999

End Function



Public Function CalcFarthestPoints(ByVal pGeometry As IGeometry, lngMethod As JenSphericalMethod, pPoint1 As IPoint, _
      pPoint2 As IPoint, dblDistance As Double, dblAZ1 As Double, dblAZ2 As Double, dblReverseAz1 As Double, _
      dblReverseAz2 As Double) As Boolean

  CalcFarthestPoints = False
      
  Dim pPtColl As IPointCollection
  Set pPtColl = pGeometry
  Dim pTestPoint1 As IPoint
  Dim pTestPoint2 As IPoint
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngPointCount As Long
  Dim dblMaxDist As Double
  Dim dblTestDist As Double
  
  Set pTestPoint1 = New Point
  Set pTestPoint2 = New Point
  
  Dim dblTestAz1 As Double
  Dim dblTestAz2 As Double
  
  Dim dblStartX As Double
  Dim dblStartY As Double
  Dim dblEndX As Double
  Dim dblEndY As Double
  Dim pClone As IClone
  
  
  dblMaxDist = -999
  lngPointCount = pPtColl.PointCount
  Debug.Print CStr(lngPointCount) & " vertices..."
  If lngPointCount > 1 Then
    For lngIndex1 = 0 To lngPointCount - 2
      pPtColl.QueryPoint lngIndex1, pTestPoint1
      
      dblStartX = pTestPoint1.x
      dblStartY = pTestPoint1.Y
      
      For lngIndex2 = lngIndex1 + 1 To lngPointCount - 1
        pPtColl.QueryPoint lngIndex2, pTestPoint2
        
        dblEndX = pTestPoint2.x
        dblEndY = pTestPoint2.Y
        
        If lngMethod = ENUM_UseSpherical Then
          dblTestDist = DistanceHaversineNumbers(dblStartY, dblStartX, dblEndY, dblEndX, , True, dblTestAz1)
        ElseIf lngMethod = ENUM_UseSpheroidal Then
          dblTestDist = DistanceVincentyNumbers2(dblStartX, dblStartY, dblEndX, dblEndY, dblTestAz1, dblTestAz2)
        Else
          dblTestDist = (((dblStartX - dblEndX) ^ 2) + ((dblStartY - dblEndY) ^ 2)) ^ (0.5)
        End If
        
'        Debug.Print "Checking [" & CStr(Format(dblStartX, "0.000")) & ", " & CStr(Format(dblStartY, "0.000")) & "] to [" & _
              CStr(Format(dblEndX, "0.000")) & ", " & CStr(Format(dblEndY, "0.000")) & "]:  Distance = " & _
              CStr(Format(dblTestDist, "0")) & " meters..."
        
        If dblTestDist > dblMaxDist Then
          
'          Debug.Print "  --> Current Shortest Distance:  " & CStr(Format(dblTestDist, "0")) & " meters..."
          
          dblMaxDist = dblTestDist
          Set pClone = pTestPoint1
          Set pPoint1 = pClone.Clone
          Set pClone = pTestPoint2
          Set pPoint2 = pClone.Clone
          
          If lngMethod = ENUM_UseSpherical Then
            dblAZ1 = dblTestAz1
            If dblAZ1 > 360 Then dblAZ1 = dblAZ1 - 360
            If dblAZ1 < 0 Then dblAZ1 = dblAZ1 + 360
            dblAZ2 = dblAZ1
          ElseIf lngMethod = ENUM_UseSpheroidal Then
            dblAZ1 = dblTestAz1
            dblAZ2 = dblTestAz2
          Else
            dblAZ1 = CalcBearingNumbers(dblStartX, dblStartY, dblEndX, dblEndY)
            If dblAZ1 > 360 Then dblAZ1 = dblAZ1 - 360
            If dblAZ1 < 0 Then dblAZ1 = dblAZ1 + 360
            dblAZ2 = dblAZ1
          End If
          
'          Debug.Print "  --> Current Shortest Distance:  " & CStr(Format(dblTestDist, "0")) & " meters..."
'          Debug.Print "  --> [" & CStr(Format(pPoint1.X, "0.000")) & ", " & CStr(Format(pPoint1.Y, "0.000")) & "] to [" & _
'              CStr(Format(pPoint2.X, "0.000")) & ", " & CStr(Format(pPoint2.Y, "0.000")) & "]"
'          Debug.Print "  --> Current Azimuth:  " & CStr(Format(dblAz1, "0")) & " degrees..."
          
        End If
        
      Next lngIndex2
    Next lngIndex1
    
    dblDistance = dblMaxDist
    
    dblReverseAz1 = dblAZ1 - 180
    If dblReverseAz1 < 0 Then dblReverseAz1 = dblReverseAz1 + 360
    dblReverseAz2 = dblAZ2 - 180
    If dblReverseAz2 < 0 Then dblReverseAz2 = dblReverseAz2 + 360
    
    CalcFarthestPoints = True
  End If

End Function

Public Function CalcBearingNumbers(dblX1 As Double, dblY1 As Double, dblX2 As Double, dblY2 As Double) As Double

  Dim dblBearing As Double

  Dim xDist As Double
  Dim yDist As Double
  Dim xyTanDeg As Double
  
  xDist = (dblX1 - dblX2)
  yDist = (dblY1 - dblY2)
  
  If xDist = 0 And yDist = 0 Then
    CalcBearingNumbers = -9999
  Else
    If yDist = 0 Then
      If xDist < 0 Then
        xyTanDeg = -90
      ElseIf xDist = 0 Then
        xyTanDeg = 0
      Else
        xyTanDeg = 90
      End If
    Else
      xyTanDeg = AsDegrees(Atn(xDist / yDist))
    End If
  
    If (yDist >= 0) Then
      dblBearing = 180 + xyTanDeg
    Else
      If (xDist <= 0) Then
        dblBearing = xyTanDeg
      Else
        dblBearing = 360 + xyTanDeg
      End If
    End If ' END CALCULATING BEARING
    
    dblBearing = Abs(dblBearing)
    CalcBearingNumbers = dblBearing
  End If

End Function

Public Function DistanceHaversine(pPointA As IPoint, pPointB As IPoint, Optional dblRadius As Double = 6371000.79000915) As Double
  
  Dim dblLat1 As Double
  Dim dblLat2 As Double
  Dim dblLat As Double
  Dim dblLong As Double
  Dim dblTemp As Double
  
  dblLat1 = DegToRad(pPointA.Y)
  dblLat2 = DegToRad(pPointB.Y)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(pPointA.x - pPointB.x)
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  DistanceHaversine = (2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))) * dblRadius

End Function

Public Function DistanceHaversineNumbers(ByVal dblLat1 As Double, ByVal dblLong1 As Double, ByVal dblLat2 As Double, _
    ByVal dblLong2 As Double, Optional dblRadius As Double = 6371000.79000915, Optional booDoAzimuth As Boolean = False, _
    Optional dblAzimuth As Double) As Double
  
  Dim dblLat As Double
  Dim dblLong As Double
  Dim dblTemp As Double
    
  dblLat1 = DegToRad(dblLat1)
  dblLat2 = DegToRad(dblLat2)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(dblLong1 - dblLong2)
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  DistanceHaversineNumbers = (2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))) * dblRadius
  
  If booDoAzimuth Then
    Dim pX As Double
    Dim QX As Double
    
    pX = DegToRad(dblLong1)
    QX = DegToRad(dblLong2)

    Dim dblTheta As Double
    Dim DeltaLong As Double
    DeltaLong = QX - pX
    dblTheta = atan2(Sin(DeltaLong) * Cos(dblLat2), Cos(dblLat1) * Sin(dblLat2) - Sin(dblLat1) * Cos(dblLat2) * Cos(DeltaLong))
    dblAzimuth = RadToDeg(dblTheta)
    If dblAzimuth < 360 Then dblAzimuth = dblAzimuth + 360
  End If

End Function
Public Function MultipointCentroidSpheroid(pMultipoint As IMultipoint, Optional dblEquatorialRadius As Double = 6378137, _
    Optional dblPolarRadius As Double = 6356752.31424518, Optional dblHeightAboveEllipsoid As Double = 0) As IPoint

  Dim pPoint As IPoint
  Dim dblX As Double
  Dim dblY As Double
  Dim dblZ As Double
  Dim dblRunningX As Double
  Dim dblRunningY As Double
  Dim dblRunningZ As Double
  Dim pPointCollection As IPointCollection
  Set pPointCollection = pMultipoint
  Dim lngIndex As Long
  dblRunningX = 0
  dblRunningY = 0
  dblRunningZ = 0
  Set pPoint = New Point
  Dim dblLongitude As Double
  Dim dblLatitude As Double
  
  Dim lngCounter As Long
  lngCounter = pPointCollection.PointCount
  For lngIndex = 0 To lngCounter - 1
    pPointCollection.QueryPoint lngIndex, pPoint
    SpheroidalLatLongToCart pPoint.x, pPoint.Y, dblX, dblY, dblZ, dblEquatorialRadius, dblPolarRadius, dblHeightAboveEllipsoid
    dblRunningX = dblRunningX + dblX
    dblRunningY = dblRunningY + dblY
    dblRunningZ = dblRunningZ + dblZ
  Next lngIndex
  
  dblX = dblRunningX / lngCounter
  dblY = dblRunningY / lngCounter
  dblZ = dblRunningZ / lngCounter
  
  ' CONVERT BACK TO GEOGRAPHIC COORDINATES
  SpheroidalCartToLatLong dblLongitude, dblLatitude, dblX, dblY, dblZ, dblEquatorialRadius, dblPolarRadius, dblHeightAboveEllipsoid
  Set MultipointCentroidSpheroid = New Point
  Set MultipointCentroidSpheroid.SpatialReference = pMultipoint.SpatialReference
  MultipointCentroidSpheroid.PutCoords dblLongitude, dblLatitude

End Function

Public Function MultipointCentroidSphere(pMultipoint As IMultipoint) As IPoint

  Dim pPoint As IPoint
  Dim dblX As Double
  Dim dblY As Double
  Dim dblZ As Double
  Dim dblRunningX As Double
  Dim dblRunningY As Double
  Dim dblRunningZ As Double
  Dim pPointCollection As IPointCollection
  Set pPointCollection = pMultipoint
  Dim lngIndex As Long
  dblRunningX = 0
  dblRunningY = 0
  dblRunningZ = 0
  Set pPoint = New Point
  Dim dblLongitude As Double
  Dim dblLatitude As Double
  
  Dim lngCounter As Long
  lngCounter = pPointCollection.PointCount
  For lngIndex = 0 To lngCounter - 1
    pPointCollection.QueryPoint lngIndex, pPoint
    SphericalLatLongToCart pPoint.x, pPoint.Y, dblX, dblY, dblZ
    dblRunningX = dblRunningX + dblX
    dblRunningY = dblRunningY + dblY
    dblRunningZ = dblRunningZ + dblZ
  Next lngIndex
  
  dblX = dblRunningX / lngCounter
  dblY = dblRunningY / lngCounter
  dblZ = dblRunningZ / lngCounter
  
  ' CONVERT BACK TO GEOGRAPHIC COORDINATES
  SphericalCartToLatLong dblLongitude, dblLatitude, dblX, dblY, dblZ
  Set MultipointCentroidSphere = New Point
  MultipointCentroidSphere.PutCoords dblLongitude, dblLatitude

End Function
Public Function MultipointCentroid(pMultipoint As IMultipoint) As IPoint

  Dim pPoint As IPoint
  Dim dblX As Double
  Dim dblY As Double
  Dim pPointCollection As IPointCollection
  Set pPointCollection = pMultipoint
  
  Dim lngIndex As Long
  dblX = 0
  dblY = 0
  Set pPoint = New Point
  For lngIndex = 0 To pPointCollection.PointCount - 1
    pPointCollection.QueryPoint lngIndex, pPoint
    dblX = dblX + pPoint.x
    dblY = dblY + pPoint.Y
  Next lngIndex
  
  Set MultipointCentroid = New Point
  MultipointCentroid.PutCoords dblX / pPointCollection.PointCount, dblY / pPointCollection.PointCount
  Set MultipointCentroid.SpatialReference = pMultipoint.SpatialReference

End Function

Public Function PolygonToPolyline(pPolygon As IPolygon2) As IPolyline
  
  Dim pNewGeometryCollection As IGeometryCollection
  Set pNewGeometryCollection = New Polyline
  Dim lngNumParts As Long
  Dim pPolyComponents() As IPolygon 'Declare an array of polygon
  Dim lngIndex As Long
  Dim pSubPolygon As IPolygon4
  Dim pTopoOp As ITopologicalOperator
  Dim pGeometryBag As IGeometryCollection
  Dim pNewSegCollection As ISegmentCollection
  Dim pOrigSegcollection As ISegmentCollection
  Dim pRing As IRing
  Dim pPath As IPath
  Dim pSpRef As ISpatialReference
  Dim lngIndex2 As Long
  Dim pPolyline As IPolyline
  
  Set pSpRef = pPolygon.SpatialReference
  Set pNewGeometryCollection = New Polyline
  
  ' GET CONNECTED COMPONENTS OF POLYGON
  lngNumParts = pPolygon.ExteriorRingCount
  ReDim pPolyComponents(lngNumParts - 1) 'Redimension the array of polygons with number of exterior rings
  pPolygon.GetConnectedComponents lngNumParts, pPolyComponents(0) 'Pass the first element of the array
  
  For lngIndex = 0 To lngNumParts - 1
    Set pSubPolygon = pPolyComponents(lngIndex)
    Set pTopoOp = pSubPolygon
    pTopoOp.Simplify
    
    ' GET SINGLE EXTERNAL RING AND ALL INTERNAL RINGS
    Set pGeometryBag = pSubPolygon.ExteriorRingBag
    Set pRing = pGeometryBag.Geometry(0)
    Set pOrigSegcollection = pRing
    Set pNewSegCollection = New Path
    pNewSegCollection.AddSegmentCollection pOrigSegcollection
    Set pPath = pNewSegCollection
    pNewGeometryCollection.AddGeometry pPath
    
    ' ADD ANY INTERNAL RINGS TO POLYLINE
    If pSubPolygon.InteriorRingCount(pRing) > 0 Then
      Set pGeometryBag = pSubPolygon.InteriorRingBag(pRing)
      For lngIndex2 = 0 To pGeometryBag.GeometryCount - 1
        Set pRing = pGeometryBag.Geometry(lngIndex2)
        Set pOrigSegcollection = pRing
        Set pNewSegCollection = New Path
        pNewSegCollection.AddSegmentCollection pOrigSegcollection
        Set pPath = pNewSegCollection
        pNewGeometryCollection.AddGeometry pPath
      Next lngIndex2
    End If
  Next lngIndex
  
  ' CLEAN NEW POLYLINE
  Set pPolyline = pNewGeometryCollection
  Set pTopoOp = pPolyline
  pTopoOp.Simplify
  Set pPolyline.SpatialReference = pSpRef

  Set PolygonToPolyline = pPolyline

End Function

Public Sub TriangleCentroid3D(dblPX As Double, dblPY As Double, dblPZ As Double, _
                                           dblQX As Double, dblQY As Double, dblQZ As Double, _
                                           dblRX As Double, dblRY As Double, dblRZ As Double, _
                                           dblCentX As Double, dblCentY As Double, dblCentZ As Double)

  dblCentX = (dblPX + dblQX + dblRX) / 3
  dblCentY = (dblPY + dblQY + dblRY) / 3
  dblCentZ = (dblPZ + dblQZ + dblRZ) / 3

End Sub

Public Sub TriangleCentroidPlane(dblPX As Double, dblPY As Double, dblQX As Double, dblQY As Double, _
                                           dblRX As Double, dblRY As Double, dblCentX As Double, dblCentY As Double)

  dblCentX = (dblPX + dblQX + dblRX) / 3
  dblCentY = (dblPY + dblQY + dblRY) / 3

End Sub


Public Sub SphericalCartToLatLong(dblLongitude As Double, dblLatitude As Double, x As Double, _
      Y As Double, Z As Double)

  ' Phi is angle from north pole down to Latitude
  ' Theta is angle from Greenwich
  
  Dim dblPhi As Double
  Dim dblTheta As Double
  
  dblPhi = atan2(Sqr(x ^ 2 + Y ^ 2), Z)
  dblTheta = atan2(Y, x)
  
  dblLongitude = RadToDeg(dblTheta)
  dblLatitude = 90 - RadToDeg(dblPhi)

End Sub


Public Sub SpheroidalCartToLatLong(dblLongitude As Double, dblLatitude As Double, x As Double, _
      Y As Double, Z As Double, Optional dblEquatorialRadius As Double = 6378137, _
      Optional dblPolarRadius As Double = 6356752.31424518, Optional dblHeightAboveEllipsoid As Double = 0)

  ' IF SPHEROID PARAMETERS NOT INCLUDED, DEFAULTS TO WGS84
  ' Phi is angle from north pole down to Latitude
  ' Theta is angle from Greenwich
  
  ' MODIFIED FROM J.C. ILIFFE, CHAPTER 2
  ' NOTE:  ILIFFE USES PHI FOR LATITUDE DIRECTLY, RATHER THAN AS DISTANCE FROM POLES
  
  Dim dblPhi As Double
  Dim dblTheta As Double
  
  
  Dim dblP As Double
  dblP = Sqr(x ^ 2 + Y ^ 2)
  
  Dim dblU As Double
  dblU = atan2((Z * dblEquatorialRadius), (dblP * dblPolarRadius))
  
  Dim dblEccentSquared As Double
  dblEccentSquared = ((dblEquatorialRadius ^ 2) - (dblPolarRadius ^ 2)) / (dblEquatorialRadius ^ 2)
  
  Dim dblEpsilon As Double
  dblEpsilon = dblEccentSquared / (1 - dblEccentSquared)
  
  dblPhi = atan2(dblP - (dblEccentSquared * dblEquatorialRadius * (Cos(dblU) ^ 3)), _
                  Z + (dblEpsilon * dblPolarRadius * (Sin(dblU) ^ 3)))
  dblTheta = atan2(Y, x)
  
  dblLongitude = RadToDeg(dblTheta)
  dblLatitude = 90 - RadToDeg(dblPhi)

End Sub

Public Sub SphericalLatLongToCart(dblLongitude As Double, dblLatitude As Double, x As Double, _
      Y As Double, Z As Double, Optional dblRadius = 1)

  ' Phi is angle from north pole down to Latitude
  ' Theta is angle from Greenwich
  
  Dim dblPhi As Double
  Dim dblTheta As Double
  
  dblPhi = DegToRad(90 - dblLatitude)
  dblTheta = DegToRad(dblLongitude)
  
  x = dblRadius * Sin(dblPhi) * Cos(dblTheta)
  Y = dblRadius * Sin(dblPhi) * Sin(dblTheta)
  Z = dblRadius * Cos(dblPhi)

End Sub


Public Sub SpheroidalLatLongToCart(dblLongitude As Double, dblLatitude As Double, x As Double, _
      Y As Double, Z As Double, Optional dblEquatorialRadius As Double = 6378137, _
      Optional dblPolarRadius As Double = 6356752.31424518, Optional dblHeightAboveEllipsoid As Double = 0)
  
  ' IF SPHEROID PARAMETERS NOT INCLUDED, DEFAULTS TO WGS84

  ' Phi is angle from north pole down to Latitude
  ' Theta is angle from Greenwich
  
  ' MODIFIED FROM J.C. ILIFFE, CHAPTER 2
  ' NOTE:  ILIFFE USES PHI FOR LATITUDE DIRECTLY, RATHER THAN AS DISTANCE FROM POLES
  
  Dim dblPhi As Double
  Dim dblTheta As Double
  
  dblPhi = DegToRad(90 - dblLatitude)
  dblTheta = DegToRad(dblLongitude)
  
  Dim dblNu As Double
  Dim dblEccentSquared As Double
  dblEccentSquared = ((dblEquatorialRadius ^ 2) - (dblPolarRadius ^ 2)) / (dblEquatorialRadius ^ 2)
  dblNu = (dblEquatorialRadius) / Sqr(1 - (dblEccentSquared * (Cos(dblPhi) ^ 2)))
  
  x = (dblNu + dblHeightAboveEllipsoid) * Sin(dblPhi) * Cos(dblTheta)
  Y = (dblNu + dblHeightAboveEllipsoid) * Sin(dblPhi) * Sin(dblTheta)
  Z = (((1 - dblEccentSquared) * dblNu) + dblHeightAboveEllipsoid) * Cos(dblPhi)

End Sub

Public Function SpheroidalPolylineMidpoint(pPolyline As IPolyline, dblDistance As Double, _
    booIsRatio As Boolean, Optional dblPolylineLength As Double) As IPoint

  Dim pPoint As IPoint
  Dim pVarArray As esriSystem.IVariantArray
  Dim pDblArray As esriSystem.IDoubleArray
  Set pVarArray = New esriSystem.varArray
  
  Dim pSegCollection As ISegmentCollection
  Dim pSeg As ISegment
  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  Dim lngIndex As Long
  Dim dblLength As Double
  Dim dblTotalLength As Double
  Dim dblAZ1 As Double
  Dim dblAZ2 As Double
  Dim dblX1 As Double
  Dim dblY1 As Double
  
  dblTotalLength = 0
  Set pSegCollection = pPolyline
  
  ' GET TOTAL LENGTH AND FILL ARRAY WITH SEGMENT STATISTICS
  For lngIndex = 0 To (pSegCollection.SegmentCount - 1)
    Set pSeg = pSegCollection.Segment(lngIndex)
    Set pPoint1 = pSeg.FromPoint
    Set pPoint2 = pSeg.ToPoint
    dblLength = DistanceVincentyPoints(pPoint1, pPoint2, dblAZ1, dblAZ2)
    dblTotalLength = dblTotalLength + dblLength
    Set pDblArray = New esriSystem.DoubleArray
    pDblArray.Add dblLength
    pDblArray.Add dblTotalLength
    pDblArray.Add dblAZ1
    pDblArray.Add pPoint1.x
    pDblArray.Add pPoint1.Y
    pVarArray.Add pDblArray
  Next lngIndex
  
  dblPolylineLength = dblTotalLength
  
  Dim dblHalfLength As Double
  If booIsRatio Then
    If dblDistance > 1 Then dblDistance = 1
    If dblDistance < 0 Then dblDistance = 0
    dblHalfLength = dblPolylineLength * dblDistance
  Else
    dblHalfLength = dblDistance
  End If
    
  For lngIndex = 0 To pVarArray.Count - 1
    Set pDblArray = pVarArray.Element(lngIndex)
    dblTotalLength = pDblArray.Element(1)
    If dblTotalLength > dblHalfLength Then
      dblLength = pDblArray.Element(0)
      dblAZ1 = pDblArray.Element(2)
      dblX1 = pDblArray.Element(3)
      dblY1 = pDblArray.Element(4)
      Exit For
    End If
  Next lngIndex
  
  Dim dblProperDistance As Double
  dblProperDistance = dblLength - (dblTotalLength - dblHalfLength)
  
  Set pPoint1 = New Point
  pPoint1.PutCoords dblX1, dblY1
  Set pPoint2 = New Point
    
  PointLineVincenty pPoint1, dblProperDistance, dblAZ1, pPoint2, dblAZ2
  
  Set SpheroidalPolylineMidpoint = pPoint2

End Function
Public Function SpheroidalPolylineMidpoint2(pPolyline As IPolyline, dblDistance As Double, _
    booIsRatio As Boolean, Optional dblPolylineLength As Double, _
    Optional dblEquatorialRadius As Double = 6378137, Optional dblPolarRadius As Double = 6356752.31424518) As IPoint
  
  ' MODIFICATION OF SpheroidalPolylineMidpoint, TO ALLOW FOR ANY ELLIPSOID
  
  Dim pPoint As IPoint
  Dim pVarArray As esriSystem.IVariantArray
  Dim pDblArray As esriSystem.IDoubleArray
  Set pVarArray = New esriSystem.varArray
  
  Dim pSegCollection As ISegmentCollection
  Dim pSeg As ISegment
  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  Dim lngIndex As Long
  Dim dblLength As Double
  Dim dblTotalLength As Double
  Dim dblAZ1 As Double
  Dim dblAZ2 As Double
  Dim dblX1 As Double
  Dim dblY1 As Double
  
  dblTotalLength = 0
  Set pSegCollection = pPolyline
  
  ' GET TOTAL LENGTH AND FILL ARRAY WITH SEGMENT STATISTICS
  For lngIndex = 0 To (pSegCollection.SegmentCount - 1)
    Set pSeg = pSegCollection.Segment(lngIndex)
    Set pPoint1 = pSeg.FromPoint
    Set pPoint2 = pSeg.ToPoint
    dblLength = DistanceVincentyPoints2(pPoint1, pPoint2, dblAZ1, dblAZ2, dblEquatorialRadius, dblPolarRadius)
    dblTotalLength = dblTotalLength + dblLength
    Set pDblArray = New esriSystem.DoubleArray
    pDblArray.Add dblLength
    pDblArray.Add dblTotalLength
    pDblArray.Add dblAZ1
    pDblArray.Add pPoint1.x
    pDblArray.Add pPoint1.Y
    pVarArray.Add pDblArray
  Next lngIndex
  
  dblPolylineLength = dblTotalLength
  
  Dim dblHalfLength As Double
  If booIsRatio Then
    If dblDistance > 1 Then dblDistance = 1
    If dblDistance < 0 Then dblDistance = 0
    dblHalfLength = dblPolylineLength * dblDistance
  Else
    dblHalfLength = dblDistance
  End If
    
  For lngIndex = 0 To pVarArray.Count - 1
    Set pDblArray = pVarArray.Element(lngIndex)
    dblTotalLength = pDblArray.Element(1)
    If dblTotalLength > dblHalfLength Then
      dblLength = pDblArray.Element(0)
      dblAZ1 = pDblArray.Element(2)
      dblX1 = pDblArray.Element(3)
      dblY1 = pDblArray.Element(4)
      Exit For
    End If
  Next lngIndex
  
  Dim dblProperDistance As Double
  dblProperDistance = dblLength - (dblTotalLength - dblHalfLength)
  
  Set pPoint1 = New Point
  pPoint1.PutCoords dblX1, dblY1
  Set pPoint1.SpatialReference = pPolyline.SpatialReference
  
  Set pPoint2 = New Point
  
  PointLineVincenty2 pPoint1, dblProperDistance, dblAZ1, pPoint2, dblAZ2
  
  Set SpheroidalPolylineMidpoint2 = pPoint2

End Function
Public Function SpheroidalPolylineLength(pPolyline As IPolyline) As Double

  ' ASSUMES POLYLINE IS IN GEOGRAPHIC COORDINATES

  Dim pSegCollection As ISegmentCollection
  Dim pSeg As ISegment
  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  Dim lngIndex As Long
  Dim dblLength As Double
  Dim dblTotalLength As Double
  Dim dblAZ1 As Double
  Dim dblAZ2 As Double
  
  dblTotalLength = 0
  Set pSegCollection = pPolyline
  
  For lngIndex = 0 To (pSegCollection.SegmentCount - 1)
    Set pSeg = pSegCollection.Segment(lngIndex)
    Set pPoint1 = pSeg.FromPoint
    Set pPoint2 = pSeg.ToPoint
    dblLength = DistanceVincentyPoints(pPoint1, pPoint2, dblAZ1, dblAZ2)
    dblTotalLength = dblTotalLength + dblLength
  Next lngIndex
  
  SpheroidalPolylineLength = dblTotalLength

End Function

Public Function SpheroidalPolylineLength2(pPolyline As IPolyline, _
    Optional dblEquatorialRadius As Double = 6378137, Optional dblPolarRadius As Double = 6356752.31424518) As Double

  ' ASSUMES POLYLINE IS IN GEOGRAPHIC COORDINATES

  Dim pSegCollection As ISegmentCollection
  Dim pSeg As ISegment
  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  Dim lngIndex As Long
  Dim dblLength As Double
  Dim dblTotalLength As Double
  Dim dblAZ1 As Double
  Dim dblAZ2 As Double
  
  dblTotalLength = 0
  Set pSegCollection = pPolyline
  
  For lngIndex = 0 To (pSegCollection.SegmentCount - 1)
    Set pSeg = pSegCollection.Segment(lngIndex)
    Set pPoint1 = pSeg.FromPoint
    Set pPoint2 = pSeg.ToPoint
    dblLength = DistanceVincentyPoints2(pPoint1, pPoint2, dblAZ1, dblAZ2, dblEquatorialRadius, dblPolarRadius)
    dblTotalLength = dblTotalLength + dblLength
  Next lngIndex
  
  SpheroidalPolylineLength2 = dblTotalLength

End Function

Public Function SphericalPolygonArea2(pPolygon As IPolygon, Optional booCalcCentroid As Boolean = False, _
      Optional dblCentX As Double, Optional dblCentY As Double, _
      Optional dblEquatorialRadius As Double = 6378137, Optional dblPolarRadius As Double = 6356752.31424518) As Double
  
  ' REMEMBER THAT THIS STILL CALCULATES ON THE SPHERE BECAUSE I HAVE NOT FIGURED OUT A WAY TO CALCULATE POLYGON AREAS
  ' ON AN ELLISOID.  THE SPHERE IS DEFINED AS THAT SPHERE WITH THE SAME VOLUME AS THE ELLIPSOID WITH THE SPECIFIED
  ' MAJOR AND MINOR AXES.
  '
  ' MODIFICATION OF SphericalPolygonArea TO ALLOW USER TO SET CUSTOM ELLIPSOID MAJOR AND MINOR AXES
  '
  ' ASSUMES POLYGON IS IN GEOGRAPHIC COORDINATES
  ' BREAK UP POLYGON INTO CONNECTED COMPONENTS
  
  Dim pPoly4 As IPolygon4
  Set pPoly4 = pPolygon
  Dim pConnected As IGeometryCollection
  Dim pRingBag As IGeometryCollection
  Dim pExtRing As IRing
  Set pConnected = pPoly4.ConnectedComponentBag
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim pPointCollection As IPointCollection
  
  Dim pArea As IArea
  Dim pCentroid As IPoint
  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  
  If booCalcCentroid Then
'    Dim dbl1Long As Double
'    Dim dbl1Lat As Double
    Dim dbl1X As Double
    Dim dbl1Y As Double
    Dim dbl1Z As Double
'    Dim dbl2Long As Double
'    Dim dbl2Lat As Double
    Dim dbl2X As Double
    Dim dbl2Y As Double
    Dim dbl2Z As Double
    Dim dbl3X As Double
    Dim dbl3Y As Double
    Dim dbl3Z As Double
    Dim dblTempCentX As Double
    Dim dblTempCentY As Double
    Dim dblTempCentZ As Double
    Dim dblRunningX As Double
    Dim dblRunningY As Double
    Dim dblRunningZ As Double
    dblRunningX = 0
    dblRunningY = 0
    dblRunningZ = 0
    Dim dblVectLength As Double
  End If
  
'  Dim dbl3Long As Double
'  Dim dbl3Lat As Double
  
  Dim pSubPolygon As IPolygon4
  Dim dblArea As Double
  Dim dblTriangleArea As Double
  dblArea = 0
  Dim dblMultiplier As Double
  Dim lngCountPos As Long
  Dim lngCountNeg As Long
  
  ' FOR TESTING CENTROID
'  Dim dblTestArea As Double
'  Dim dblTestRunningArea As Double
'  Dim dblTestX As Double
'  Dim dblTestY As Double
'  Dim dblTestRunningX As Double
'  Dim dblTestRunningY As Double
  
  Dim pSegCollection As ISegmentCollection
  Dim pSeg As ISegment
  For lngIndex = 0 To (pConnected.GeometryCount - 1)
    Set pSubPolygon = pConnected.Geometry(lngIndex)
    Set pArea = pSubPolygon
    Set pCentroid = pArea.Centroid
    Set pSegCollection = pSubPolygon
    
'    dbl3Long = pCentroid.X
'    dbl3Lat = pCentroid.Y
    
    For lngIndex2 = 0 To (pSegCollection.SegmentCount - 1)
      Set pSeg = pSegCollection.Segment(lngIndex2)
      Set pPoint1 = pSeg.FromPoint
      Set pPoint2 = pSeg.ToPoint
      
'      ThisDocument.Graphic_MakeFromGeometry Document, pPoint1, "delete_corridors"
'      ThisDocument.Graphic_MakeFromGeometry Document, pPoint2, "delete_corridors"
      dblTriangleArea = SphericalTriangleArea2(pPoint1, pPoint2, pCentroid, dblMultiplier, dblEquatorialRadius, dblPolarRadius)
      dblArea = dblArea + dblTriangleArea
      
      ' FOR TESTING CENTROID
'      dblTestArea = TriangleAreaPoints3DValues(pPoint1.X, pPoint1.Y, 1, pPoint2.X, pPoint2.Y, _
'                1, pCentroid.X, pCentroid.Y, 1)
'      dblTestRunningArea = dblTestRunningArea + dblTestArea
'      TriangleCentroidPlane pPoint1.X, pPoint1.Y, pPoint2.X, pPoint2.Y, pCentroid.X, pCentroid.Y, dblTestX, dblTestY
'      dblTestRunningX = dblTestRunningX + (dblTestX * dblTestArea)
'      dblTestRunningY = dblTestRunningY + (dblTestY * dblTestArea)
      
      If booCalcCentroid Then
        SphericalLatLongToCart pPoint1.x, pPoint1.Y, dbl1X, dbl1Y, dbl1Z
        SphericalLatLongToCart pPoint2.x, pPoint2.Y, dbl2X, dbl2Y, dbl2Z
        SphericalLatLongToCart pCentroid.x, pCentroid.Y, dbl3X, dbl3Y, dbl3Z
        TriangleCentroid3D dbl1X, dbl1Y, dbl1Z, dbl2X, dbl2Y, dbl2Z, dbl3X, dbl3Y, dbl3Z, _
                dblTempCentX, dblTempCentY, dblTempCentZ
    
        ' NORMALIZE VECTOR
        dblVectLength = Sqr(dblTempCentX ^ 2 + dblTempCentY ^ 2 + dblTempCentZ ^ 2)
        dblTempCentX = dblTempCentX / dblVectLength
        dblTempCentY = dblTempCentY / dblVectLength
        dblTempCentZ = dblTempCentZ / dblVectLength
        
        dblRunningX = dblRunningX + (dblTempCentX * dblTriangleArea)
        dblRunningY = dblRunningY + (dblTempCentY * dblTriangleArea)
        dblRunningZ = dblRunningZ + (dblTempCentZ * dblTriangleArea)
      End If
    Next lngIndex2
  Next lngIndex
  
  If booCalcCentroid Then
    If dblArea > 0 Then
      dblRunningX = dblRunningX / dblArea
      dblRunningY = dblRunningY / dblArea
      dblRunningZ = dblRunningZ / dblArea
      
      SphericalCartToLatLong dblCentX, dblCentY, dblRunningX, dblRunningY, dblRunningZ
    Else
      
      ' IF AREA = 0 BUT HAS VERTICES, THEN CALCULATE CENTROID AS MULTIPOINT CENTROID?  NO; MIGHT HAVE OVERLAPPING VERTICES THAT WOULD SKEW
      ' IF HAS NO MASS, THEN CANNOT HAVE CENTER OF MASS
      
      dblCentX = -9999
      dblCentY = -9999
    End If
  End If
  
  ' FOR TESTING CENTROID
'  dblTestRunningX = dblTestRunningX / dblTestRunningArea
'  dblTestRunningY = dblTestRunningY / dblTestRunningArea
'  Debug.Print "Test Centroid:  X = " & dblTestRunningX & ",  Y = " & dblTestRunningY
  
  SphericalPolygonArea2 = dblArea

End Function
Public Function SphericalPolygonArea(pPolygon As IPolygon, Optional booCalcCentroid As Boolean = False, _
      Optional dblCentX As Double, Optional dblCentY As Double) As Double

  ' ASSUMES POLYGON IS IN GEOGRAPHIC COORDINATES
  ' BREAK UP POLYGON INTO CONNECTED COMPONENTS
  Dim pPoly4 As IPolygon4
  Set pPoly4 = pPolygon
  Dim pConnected As IGeometryCollection
  Dim pRingBag As IGeometryCollection
  Dim pExtRing As IRing
  Set pConnected = pPoly4.ConnectedComponentBag
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  Dim pPointCollection As IPointCollection
  
  Dim pArea As IArea
  Dim pCentroid As IPoint
  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  
  If booCalcCentroid Then
'    Dim dbl1Long As Double
'    Dim dbl1Lat As Double
    Dim dbl1X As Double
    Dim dbl1Y As Double
    Dim dbl1Z As Double
'    Dim dbl2Long As Double
'    Dim dbl2Lat As Double
    Dim dbl2X As Double
    Dim dbl2Y As Double
    Dim dbl2Z As Double
    Dim dbl3X As Double
    Dim dbl3Y As Double
    Dim dbl3Z As Double
    Dim dblTempCentX As Double
    Dim dblTempCentY As Double
    Dim dblTempCentZ As Double
    Dim dblRunningX As Double
    Dim dblRunningY As Double
    Dim dblRunningZ As Double
    dblRunningX = 0
    dblRunningY = 0
    dblRunningZ = 0
    Dim dblVectLength As Double
  End If
  
'  Dim dbl3Long As Double
'  Dim dbl3Lat As Double
  
  Dim pSubPolygon As IPolygon4
  Dim dblArea As Double
  Dim dblTriangleArea As Double
  dblArea = 0
  Dim dblMultiplier As Double
  Dim lngCountPos As Long
  Dim lngCountNeg As Long
  
  ' FOR TESTING CENTROID
'  Dim dblTestArea As Double
'  Dim dblTestRunningArea As Double
'  Dim dblTestX As Double
'  Dim dblTestY As Double
'  Dim dblTestRunningX As Double
'  Dim dblTestRunningY As Double
  
  Dim pSegCollection As ISegmentCollection
  Dim pSeg As ISegment
  For lngIndex = 0 To (pConnected.GeometryCount - 1)
    Set pSubPolygon = pConnected.Geometry(lngIndex)
    Set pArea = pSubPolygon
    Set pCentroid = pArea.Centroid
    Set pSegCollection = pSubPolygon
    
'    dbl3Long = pCentroid.X
'    dbl3Lat = pCentroid.Y
    
    For lngIndex2 = 0 To (pSegCollection.SegmentCount - 1)
      Set pSeg = pSegCollection.Segment(lngIndex2)
      Set pPoint1 = pSeg.FromPoint
      Set pPoint2 = pSeg.ToPoint
      
'      ThisDocument.Graphic_MakeFromGeometry Document, pPoint1, "delete_corridors"
'      ThisDocument.Graphic_MakeFromGeometry Document, pPoint2, "delete_corridors"
      dblTriangleArea = SphericalTriangleArea(pPoint1, pPoint2, pCentroid, dblMultiplier)
      dblArea = dblArea + dblTriangleArea
      
      ' FOR TESTING CENTROID
'      dblTestArea = TriangleAreaPoints3DValues(pPoint1.X, pPoint1.Y, 1, pPoint2.X, pPoint2.Y, _
'                1, pCentroid.X, pCentroid.Y, 1)
'      dblTestRunningArea = dblTestRunningArea + dblTestArea
'      TriangleCentroidPlane pPoint1.X, pPoint1.Y, pPoint2.X, pPoint2.Y, pCentroid.X, pCentroid.Y, dblTestX, dblTestY
'      dblTestRunningX = dblTestRunningX + (dblTestX * dblTestArea)
'      dblTestRunningY = dblTestRunningY + (dblTestY * dblTestArea)
      
      If booCalcCentroid Then
        SphericalLatLongToCart pPoint1.x, pPoint1.Y, dbl1X, dbl1Y, dbl1Z
        SphericalLatLongToCart pPoint2.x, pPoint2.Y, dbl2X, dbl2Y, dbl2Z
        SphericalLatLongToCart pCentroid.x, pCentroid.Y, dbl3X, dbl3Y, dbl3Z
        TriangleCentroid3D dbl1X, dbl1Y, dbl1Z, dbl2X, dbl2Y, dbl2Z, dbl3X, dbl3Y, dbl3Z, _
                dblTempCentX, dblTempCentY, dblTempCentZ
    
        ' NORMALIZE VECTOR
        dblVectLength = Sqr(dblTempCentX ^ 2 + dblTempCentY ^ 2 + dblTempCentZ ^ 2)
        dblTempCentX = dblTempCentX / dblVectLength
        dblTempCentY = dblTempCentY / dblVectLength
        dblTempCentZ = dblTempCentZ / dblVectLength
        
        dblRunningX = dblRunningX + (dblTempCentX * dblTriangleArea)
        dblRunningY = dblRunningY + (dblTempCentY * dblTriangleArea)
        dblRunningZ = dblRunningZ + (dblTempCentZ * dblTriangleArea)
      End If
    Next lngIndex2
  Next lngIndex
  
  If booCalcCentroid Then
    If dblArea > 0 Then
      dblRunningX = dblRunningX / dblArea
      dblRunningY = dblRunningY / dblArea
      dblRunningZ = dblRunningZ / dblArea
      
      SphericalCartToLatLong dblCentX, dblCentY, dblRunningX, dblRunningY, dblRunningZ
    Else
      
      ' IF AREA = 0 BUT HAS VERTICES, THEN CALCULATE CENTROID AS MULTIPOINT CENTROID?  NO; MIGHT HAVE OVERLAPPING VERTICES THAT WOULD SKEW
      ' IF HAS NO MASS, THEN CANNOT HAVE CENTER OF MASS
      
      dblCentX = -9999
      dblCentY = -9999
    End If
  End If
  
  ' FOR TESTING CENTROID
'  dblTestRunningX = dblTestRunningX / dblTestRunningArea
'  dblTestRunningY = dblTestRunningY / dblTestRunningArea
'  Debug.Print "Test Centroid:  X = " & dblTestRunningX & ",  Y = " & dblTestRunningY
  
  SphericalPolygonArea = dblArea

End Function

Public Function AzimuthHaversine(pPointA As IPoint, pPointB As IPoint) As Double

  ' WITH HELP FROM http://www.movable-type.co.uk/scripts/latlong.html
  ' Formula:    Theta = atan2( sin(Deltalong)*cos(lat2), cos(lat1)*sin(lat2) - sin(lat1)*cos(lat2)*cos(DeltaLong) )
  
  Dim DeltaLong As Double
  Dim pX As Double
  Dim PY As Double
  Dim QX As Double
  Dim QY As Double
  
  pX = DegToRad(pPointA.x)
  PY = DegToRad(pPointA.Y)
  QX = DegToRad(pPointB.x)
  QY = DegToRad(pPointB.Y)
  
  DeltaLong = QX - pX
  Dim dblTheta As Double
  dblTheta = atan2(Sin(DeltaLong) * Cos(QY), Cos(PY) * Sin(QY) - Sin(PY) * Cos(QY) * Cos(DeltaLong))
  AzimuthHaversine = RadToDeg(dblTheta)
  If AzimuthHaversine < 0 Then AzimuthHaversine = AzimuthHaversine + 360
  If AzimuthHaversine > 360 Then AzimuthHaversine = AzimuthHaversine - 360

End Function

Public Function SphericalTriangleArea2(pPointA As IPoint, pPointB As IPoint, pPointC As IPoint, Optional dblMult As Double, _
    Optional dblEquatorialRadius As Double = 6378137, Optional dblPolarRadius As Double = 6356752.31424518) As Double

  ' MODIFICATION OF SphericalTriangleArea, TO ALLOW USER TO OPTIONALLY SEND CUSTOM MAJOR AND MINOR ELLIPSOID AXES
  
  ' BASED ON GIRARD'S FORMULA:  Area = R^2 * (A + B + C - Pi)
  '                          Where A = Angle 1
  '                                B = Angle 2
  '                                C = Angle 3
  '                   A + B + C - Pi = Spherical Excess
  '                                R = Sphere Radius
  ' Trick is to get Angles A, B and C from points.
  '
  ' ANOTHER FORMULATION, BASED ON DISTANCES:
  '                       Tan(E / 4) = sqrt(Tan(S / 2) * Tan((S - A) / 2) * Tan((S - B) / 2) * Tan((S - C) / 2))
  '                 Spherical Excess = E
  '                   where  a, b, c = sides of spherical triangle
  '                                S = (A + B + C) / 2
  ' INITAL AZIMUTH = atn( sin (Lo2 - Lo1) / (cos (Lo2 - Lo1) sin L1 - cos L1 tan L2)
  '         http://fer3.com/arc/m2.aspx?i=1688&y=200111
  
  ' FOR DEBUGGING
'  Static dblMax As Double
'  Static dblMin As Double
  
  If Abs(pPointA.x - pPointB.x) < 0.000000000001 And Abs(pPointA.x - pPointC.x) < 0.000000000001 Then
    SphericalTriangleArea2 = 0
    Exit Function
  End If
  
  If Abs(pPointA.x - pPointB.x) < 0.000000000001 And Abs(pPointA.Y - pPointB.Y) < 0.000000000001 Then
    SphericalTriangleArea2 = 0
    Exit Function
  End If
  
  If Abs(pPointA.x - pPointC.x) < 0.000000000001 And Abs(pPointA.Y - pPointC.Y) < 0.000000000001 Then
    SphericalTriangleArea2 = 0
    Exit Function
  End If
  
  If Abs(pPointB.x - pPointC.x) < 0.000000000001 And Abs(pPointB.Y - pPointC.Y) < 0.000000000001 Then
    SphericalTriangleArea2 = 0
    Exit Function
  End If
  
  ' SPECIAL CASE IF TWO POINTS AT POLE
  Dim lngPoleCounter As Long
  lngPoleCounter = 0
  If Abs(Abs(pPointA.Y) - 90) < 0.000000001 Then lngPoleCounter = lngPoleCounter + 1
  If Abs(Abs(pPointB.Y) - 90) < 0.000000001 Then lngPoleCounter = lngPoleCounter + 1
  If Abs(Abs(pPointC.Y) - 90) < 0.000000001 Then lngPoleCounter = lngPoleCounter + 1
  If lngPoleCounter > 1 Then
    SphericalTriangleArea2 = 0
    Exit Function
  End If
  
  Dim dblAB As Double
  Dim dblBC As Double
  Dim dblCA As Double
  
  Dim dblR As Double         ' RADIUS
  dblR = (dblEquatorialRadius ^ 2 * dblPolarRadius) ^ (1 / 3)   ' PROPER 3-AXIS GEOMETRIC MEAN; (a^2 * b) ^ (1/3)
  
  Dim dblLat As Double
  Dim dblLong As Double
  Dim dblTemp As Double
  Dim dblLat1 As Double
  Dim dblLat2 As Double
  Dim dblC As Double
  Dim dblAzimuthAB As Double
  Dim dblAzimuthBC As Double
  Dim dblAzimuthAC As Double
  Dim dblAzimuthAA As Double
  Dim dblAzRev As Double
  Dim dblMultiplier As Double
  Dim dblLong2 As Double
  
  ' CALCULATE LENGTH OF GEOCURVE AB USING HAVERSINE FORMULA
  dblLat1 = DegToRad(pPointA.Y)
  dblLat2 = DegToRad(pPointB.Y)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(pPointA.x - pPointB.x)
  dblLong2 = -dblLong
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  dblAB = 2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))
  ' AZIMUTH FROM A TO B
'  DistanceVincentyNumbers pPointA.X, pPointA.Y, pPointB.X, pPointB.Y, dblAzimuthAB, dblAzRev
'  dblAB = DistanceVincentyNumbers(pPointA.x, pPointA.y, pPointB.x, pPointB.y, dblAzimuthAB, dblAzRev) / dblR
'  Debug.Print "A to B:  Vincenty = " & dblAzimuthAB
  dblAzimuthAB = atan2(Sin(-dblLong) * Cos(dblLat2), _
        Cos(dblLat1) * Sin(dblLat2) - Sin(dblLat1) * Cos(dblLat2) * Cos(-dblLong))
'  Debug.Print "A to B:  Simpler = " & dblAzimuthAB
'  dblAB = dblR * dblC
  
  ' CALCULATE LENGTH OF GEOCURVE BC USING HAVERSINE FORMULA
  dblLat1 = DegToRad(pPointB.Y)
  dblLat2 = DegToRad(pPointC.Y)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(pPointB.x - pPointC.x)
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  dblBC = 2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))
'  dblBC = DistanceVincentyNumbers(pPointB.x, pPointB.y, pPointC.x, pPointC.y, dblAzimuthBC, dblAzRev) / dblR
'  dblBC = DistanceVincentyNumbers(pPointB.x, pPointB.y, pPointC.x, pPointC.y, dblAzimuthBC, dblAzRev) / dblR
'  dblBC = dblR * dblC
  
  ' CALCULATE LENGTH OF GEOCURVE AB USING HAVERSINE FORMULA
  dblLat1 = DegToRad(pPointC.Y)
  dblLat2 = DegToRad(pPointA.Y)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(pPointC.x - pPointA.x)
  
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  dblCA = 2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))
  ' AZIMUTH FROM A TO C
'  DistanceVincentyNumbers pPointA.x, pPointA.y, pPointC.x, pPointC.y, dblAzimuthAC, dblAzRev
'  dblCA = DistanceVincentyNumbers(pPointA.x, pPointA.y, pPointC.x, pPointC.y, dblAzimuthAC, dblAzRev) / dblR
'  Debug.Print "A to C:  Vincenty = " & dblAzimuth
'  dblAzimuth2 = RadToDeg(atan2(Sin(pPointC.x - pPointA.x), _
                    (Cos(pPointC.x - pPointA.x) * Sin(pPointA.y) - Cos(pPointA.y) * Tan(pPointC.y))))
'  dblCA = dblR * dblC
'  Debug.Print "Az1 = " & CStr(dblAzimuth1) & ",       Az2 = " & CStr(dblAzimuth2)

  ' NOTE:  Lat1 and Lat2 flipped in equation below because variables originally defined for Line CA, not Line AC
  dblAzimuthAC = atan2(Sin(dblLong) * Cos(dblLat1), _
        Cos(dblLat2) * Sin(dblLat1) - Sin(dblLat2) * Cos(dblLat1) * Cos(dblLong))
'  Debug.Print "A to C:  Simpler = " & dblAzimuthAC
  
'  If dblAzimuthAB > dblMax Then dblMax = dblAzimuthAB
'  If dblAzimuthAB < dblMin Then dblMin = dblAzimuthAB
'  Debug.Print "Current Max = " & dblMax & ",   Current Min = " & dblMin
  
  If dblAzimuthAB < 0 Then dblAzimuthAB = dblAzimuthAB + 2 * dblPI
  If dblAzimuthAC < 0 Then dblAzimuthAC = dblAzimuthAC + 2 * dblPI
  
  Dim dblDiff As Double
  dblDiff = dblAzimuthAC - dblAzimuthAB
  If dblDiff > 0 Then              ' EITHER AC > AB or AC IS TO THE LEFT OF NORTH
    If dblDiff > dblPI Then         ' THEN AC IS TO THE LEFT OF NORTH
      dblMultiplier = -1           ' COUNTERCLOCKWISE
    Else                           ' THEN AC > AB
      dblMultiplier = 1            ' CLOCKWISE
    End If
  Else                             ' EITHER AC < AB or AB IS TO THE LEFT OF NORTH
    If Abs(dblDiff) > dblPI Then   ' THEN AB IS TO THE LEFT OF NORTH
      dblMultiplier = 1            ' CLOCKWISE
    Else                           ' THEN AC < AB
      dblMultiplier = -1           ' COUNTERCLOCKWISE
    End If
  End If
    
  
'  If ((dblDiff > dblPi) And (dblAzimuthAB > dblAzimuthAC)) Or ((dblDiff < dblPi) And (dblAzimuthAB < dblAzimuthAC)) Then
'    ' IS CLOCKWISE
'    dblMultiplier = 1
'  Else
'    ' IS COUNTERCLOCKWISE
'    dblMultiplier = -1
'  End If

'  If Abs(dblAzimuthAC - dblAzimuthAB) > dblPi Then dblAzimuthAB = dblAzimuth
'  dblAzimuthAB = dblAzimuthAB + 360
'  dblAzimuthAC = dblAzimuthAC + 360
'  If dblAzimuthAB > dblAzimuthAC Then
'    dblMultiplier = -1
'  Else
'    dblMultiplier = 1
'  End If
  
'  Dim booIsClockwise As Boolean
'  booIsClockwise = CalcCheckClockwise(pPointA, pPointB, pPointC)
'  If booIsClockwise Then
'    If dblMultiplier = -1 Then
'      Debug.Print "-1, 1:  AB = " & CStr(RadToDeg(dblAzimuthAB)) & ",   AC = " & CStr(RadToDeg(dblAzimuthAC)) & vbCrLf & _
'                  "  --> Point A:  X = " & pPointA.x & ",    Y = " & pPointA.y & vbCrLf & _
'                  "  --> Point B:  X = " & pPointB.x & ",    Y = " & pPointB.y & vbCrLf & _
'                  "  --> Point C:  X = " & pPointC.x & ",    Y = " & pPointC.y
'    End If
'  Else
'    If dblMultiplier = 1 Then
'      Debug.Print "1, -1:  AB = " & CStr(RadToDeg(dblAzimuthAB)) & ",  AC = " & CStr(RadToDeg(dblAzimuthAC)) & vbCrLf & _
'                  "  --> Point A:  X = " & pPointA.x & ",    Y = " & pPointA.y & vbCrLf & _
'                  "  --> Point B:  X = " & pPointB.x & ",    Y = " & pPointB.y & vbCrLf & _
'                  "  --> Point C:  X = " & pPointC.x & ",    Y = " & pPointC.y
'    End If
'  End If
  
  Dim dblS As Double
  dblS = (dblAB + dblBC + dblCA) / 2
  
  Dim dblTan_S_AB As Double
  Dim dblTan_S_BC As Double
  Dim dblTan_S_AC As Double
  
  dblTan_S_AB = Tan((dblS - dblAB) / 2)
  dblTan_S_BC = Tan((dblS - dblBC) / 2)
  dblTan_S_AC = Tan((dblS - dblCA) / 2)
  
  If dblS < 0 Then dblS = 0
  If dblTan_S_AB < 0 Then dblTan_S_AB = 0
  If dblTan_S_BC < 0 Then dblTan_S_BC = 0
  If dblTan_S_AC < 0 Then dblTan_S_AC = 0
  
  Dim dblTanEOver4 As Double
'  dblTanEOver4 = Sqr(Tan(dblS / 2) * Tan((dblS - dblAB) / 2) * Tan((dblS - dblBC) / 2) * Tan((dblS - dblCA) / 2))
  dblTanEOver4 = Sqr(Tan(dblS / 2) * dblTan_S_AB * dblTan_S_BC * dblTan_S_AC)
  
'  Dim dblTanEOver4 As Double
'  dblTanEOver4 = Sqr(Tan(dblS / 2) * Tan((dblS - dblAB) / 2) * Tan((dblS - dblBC) / 2) * Tan((dblS - dblCA) / 2))
  Dim dblE As Double
  dblE = Atn(dblTanEOver4) * 4
  
  dblMult = dblMultiplier
  SphericalTriangleArea2 = dblR ^ 2 * dblE * dblMultiplier

End Function


Public Function SphericalTriangleArea(pPointA As IPoint, pPointB As IPoint, pPointC As IPoint, Optional dblMult As Double) As Double
  
  
  ' BASED ON GIRARD'S FORMULA:  Area = R^2 * (A + B + C - Pi)
  '                          Where A = Angle 1
  '                                B = Angle 2
  '                                C = Angle 3
  '                   A + B + C - Pi = Spherical Excess
  '                                R = Sphere Radius
  ' Trick is to get Angles A, B and C from points.
  '
  ' ANOTHER FORMULATION, BASED ON DISTANCES:
  '                       Tan(E / 4) = sqrt(Tan(S / 2) * Tan((S - A) / 2) * Tan((S - B) / 2) * Tan((S - C) / 2))
  '                 Spherical Excess = E
  '                   where  a, b, c = sides of spherical triangle
  '                                S = (A + B + C) / 2
  ' INITAL AZIMUTH = atn( sin (Lo2 - Lo1) / (cos (Lo2 - Lo1) sin L1 - cos L1 tan L2)
  '         http://fer3.com/arc/m2.aspx?i=1688&y=200111
  
  ' FOR DEBUGGING
'  Static dblMax As Double
'  Static dblMin As Double
  
  If Abs(pPointA.x - pPointB.x) < 0.000000000001 And Abs(pPointA.x - pPointC.x) < 0.000000000001 Then
    SphericalTriangleArea = 0
    Exit Function
  End If
  
  If Abs(pPointA.x - pPointB.x) < 0.000000000001 And Abs(pPointA.Y - pPointB.Y) < 0.000000000001 Then
    SphericalTriangleArea = 0
    Exit Function
  End If
  
  If Abs(pPointA.x - pPointC.x) < 0.000000000001 And Abs(pPointA.Y - pPointC.Y) < 0.000000000001 Then
    SphericalTriangleArea = 0
    Exit Function
  End If
  
  If Abs(pPointB.x - pPointC.x) < 0.000000000001 And Abs(pPointB.Y - pPointC.Y) < 0.000000000001 Then
    SphericalTriangleArea = 0
    Exit Function
  End If
  
  ' SPECIAL CASE IF TWO POINTS AT POLE
  Dim lngPoleCounter As Long
  lngPoleCounter = 0
  If Abs(Abs(pPointA.Y) - 90) < 0.000000001 Then lngPoleCounter = lngPoleCounter + 1
  If Abs(Abs(pPointB.Y) - 90) < 0.000000001 Then lngPoleCounter = lngPoleCounter + 1
  If Abs(Abs(pPointC.Y) - 90) < 0.000000001 Then lngPoleCounter = lngPoleCounter + 1
  If lngPoleCounter > 1 Then
    SphericalTriangleArea = 0
    Exit Function
  End If
  
  Dim dblAB As Double
  Dim dblBC As Double
  Dim dblCA As Double
  
  Dim dblR As Double         ' RADIUS
'  dblR = (6378137 + 6356752.31424518) / 2 ' AVERAGE OF WGS ELLIPSOID MAJOR AND MINOR AXES
'  dblR = Sqr(6378137 * 6356752.31424518)  ' GEOMETRIC MEAN OF WGS ELLIPSOID MAJOR AND MINOR AXES
  dblR = (6378137 ^ 2 * 6356752.31424518) ^ (1 / 3)   ' PROPER 3-AXIS GEOMETRIC MEAN; (a^2 * b) ^ (1/3)
  
  Dim dblLat As Double
  Dim dblLong As Double
  Dim dblTemp As Double
  Dim dblLat1 As Double
  Dim dblLat2 As Double
  Dim dblC As Double
  Dim dblAzimuthAB As Double
  Dim dblAzimuthBC As Double
  Dim dblAzimuthAC As Double
  Dim dblAzimuthAA As Double
  Dim dblAzRev As Double
  Dim dblMultiplier As Double
  Dim dblLong2 As Double
  
  ' CALCULATE LENGTH OF GEOCURVE AB USING HAVERSINE FORMULA
  dblLat1 = DegToRad(pPointA.Y)
  dblLat2 = DegToRad(pPointB.Y)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(pPointA.x - pPointB.x)
  dblLong2 = -dblLong
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  dblAB = 2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))
  ' AZIMUTH FROM A TO B
'  DistanceVincentyNumbers pPointA.X, pPointA.Y, pPointB.X, pPointB.Y, dblAzimuthAB, dblAzRev
'  dblAB = DistanceVincentyNumbers(pPointA.x, pPointA.y, pPointB.x, pPointB.y, dblAzimuthAB, dblAzRev) / dblR
'  Debug.Print "A to B:  Vincenty = " & dblAzimuthAB
  dblAzimuthAB = atan2(Sin(-dblLong) * Cos(dblLat2), _
        Cos(dblLat1) * Sin(dblLat2) - Sin(dblLat1) * Cos(dblLat2) * Cos(-dblLong))
'  Debug.Print "A to B:  Simpler = " & dblAzimuthAB
'  dblAB = dblR * dblC
  
  ' CALCULATE LENGTH OF GEOCURVE BC USING HAVERSINE FORMULA
  dblLat1 = DegToRad(pPointB.Y)
  dblLat2 = DegToRad(pPointC.Y)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(pPointB.x - pPointC.x)
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  dblBC = 2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))
'  dblBC = DistanceVincentyNumbers(pPointB.x, pPointB.y, pPointC.x, pPointC.y, dblAzimuthBC, dblAzRev) / dblR
'  dblBC = DistanceVincentyNumbers(pPointB.x, pPointB.y, pPointC.x, pPointC.y, dblAzimuthBC, dblAzRev) / dblR
'  dblBC = dblR * dblC
  
  ' CALCULATE LENGTH OF GEOCURVE AB USING HAVERSINE FORMULA
  dblLat1 = DegToRad(pPointC.Y)
  dblLat2 = DegToRad(pPointA.Y)
  dblLat = dblLat1 - dblLat2
  dblLong = DegToRad(pPointC.x - pPointA.x)
  
  dblTemp = (Sin(dblLat / 2)) ^ 2 + Cos(dblLat1) * Cos(dblLat2) * (Sin(dblLong / 2)) ^ 2
  dblCA = 2 * atan2(Sqr(dblTemp), Sqr(1 - dblTemp))
  ' AZIMUTH FROM A TO C
'  DistanceVincentyNumbers pPointA.x, pPointA.y, pPointC.x, pPointC.y, dblAzimuthAC, dblAzRev
'  dblCA = DistanceVincentyNumbers(pPointA.x, pPointA.y, pPointC.x, pPointC.y, dblAzimuthAC, dblAzRev) / dblR
'  Debug.Print "A to C:  Vincenty = " & dblAzimuth
'  dblAzimuth2 = RadToDeg(atan2(Sin(pPointC.x - pPointA.x), _
                    (Cos(pPointC.x - pPointA.x) * Sin(pPointA.y) - Cos(pPointA.y) * Tan(pPointC.y))))
'  dblCA = dblR * dblC
'  Debug.Print "Az1 = " & CStr(dblAzimuth1) & ",       Az2 = " & CStr(dblAzimuth2)

  ' NOTE:  Lat1 and Lat2 flipped in equation below because variables originally defined for Line CA, not Line AC
  dblAzimuthAC = atan2(Sin(dblLong) * Cos(dblLat1), _
        Cos(dblLat2) * Sin(dblLat1) - Sin(dblLat2) * Cos(dblLat1) * Cos(dblLong))
'  Debug.Print "A to C:  Simpler = " & dblAzimuthAC
  
'  If dblAzimuthAB > dblMax Then dblMax = dblAzimuthAB
'  If dblAzimuthAB < dblMin Then dblMin = dblAzimuthAB
'  Debug.Print "Current Max = " & dblMax & ",   Current Min = " & dblMin
  
  If dblAzimuthAB < 0 Then dblAzimuthAB = dblAzimuthAB + 2 * dblPI
  If dblAzimuthAC < 0 Then dblAzimuthAC = dblAzimuthAC + 2 * dblPI
  
  Dim dblDiff As Double
  dblDiff = dblAzimuthAC - dblAzimuthAB
  If dblDiff > 0 Then              ' EITHER AC > AB or AC IS TO THE LEFT OF NORTH
    If dblDiff > dblPI Then         ' THEN AC IS TO THE LEFT OF NORTH
      dblMultiplier = -1           ' COUNTERCLOCKWISE
    Else                           ' THEN AC > AB
      dblMultiplier = 1            ' CLOCKWISE
    End If
  Else                             ' EITHER AC < AB or AB IS TO THE LEFT OF NORTH
    If Abs(dblDiff) > dblPI Then   ' THEN AB IS TO THE LEFT OF NORTH
      dblMultiplier = 1            ' CLOCKWISE
    Else                           ' THEN AC < AB
      dblMultiplier = -1           ' COUNTERCLOCKWISE
    End If
  End If
    
  
'  If ((dblDiff > dblPi) And (dblAzimuthAB > dblAzimuthAC)) Or ((dblDiff < dblPi) And (dblAzimuthAB < dblAzimuthAC)) Then
'    ' IS CLOCKWISE
'    dblMultiplier = 1
'  Else
'    ' IS COUNTERCLOCKWISE
'    dblMultiplier = -1
'  End If

'  If Abs(dblAzimuthAC - dblAzimuthAB) > dblPi Then dblAzimuthAB = dblAzimuth
'  dblAzimuthAB = dblAzimuthAB + 360
'  dblAzimuthAC = dblAzimuthAC + 360
'  If dblAzimuthAB > dblAzimuthAC Then
'    dblMultiplier = -1
'  Else
'    dblMultiplier = 1
'  End If
  
'  Dim booIsClockwise As Boolean
'  booIsClockwise = CalcCheckClockwise(pPointA, pPointB, pPointC)
'  If booIsClockwise Then
'    If dblMultiplier = -1 Then
'      Debug.Print "-1, 1:  AB = " & CStr(RadToDeg(dblAzimuthAB)) & ",   AC = " & CStr(RadToDeg(dblAzimuthAC)) & vbCrLf & _
'                  "  --> Point A:  X = " & pPointA.x & ",    Y = " & pPointA.y & vbCrLf & _
'                  "  --> Point B:  X = " & pPointB.x & ",    Y = " & pPointB.y & vbCrLf & _
'                  "  --> Point C:  X = " & pPointC.x & ",    Y = " & pPointC.y
'    End If
'  Else
'    If dblMultiplier = 1 Then
'      Debug.Print "1, -1:  AB = " & CStr(RadToDeg(dblAzimuthAB)) & ",  AC = " & CStr(RadToDeg(dblAzimuthAC)) & vbCrLf & _
'                  "  --> Point A:  X = " & pPointA.x & ",    Y = " & pPointA.y & vbCrLf & _
'                  "  --> Point B:  X = " & pPointB.x & ",    Y = " & pPointB.y & vbCrLf & _
'                  "  --> Point C:  X = " & pPointC.x & ",    Y = " & pPointC.y
'    End If
'  End If
  
  Dim dblS As Double
  dblS = (dblAB + dblBC + dblCA) / 2
  
  Dim dblTanEOver4 As Double
  dblTanEOver4 = Sqr(Tan(dblS / 2) * Tan((dblS - dblAB) / 2) * Tan((dblS - dblBC) / 2) * Tan((dblS - dblCA) / 2))
  Dim dblE As Double
  dblE = Atn(dblTanEOver4) * 4
  
  dblMult = dblMultiplier
  SphericalTriangleArea = dblR ^ 2 * dblE * dblMultiplier

End Function

Public Function ArcSinJen(dblValue As Double) As Double

'  ArcSinJen = Atn(dblValue / Sqr(-dblValue * dblValue + 1))
  
  If dblValue = 1 Then
    ArcSinJen = dblPI / 2
  ElseIf dblValue = -1 Then
    ArcSinJen = -dblPI / 2
  Else
    ArcSinJen = Atn(dblValue / Sqr(-dblValue * dblValue + 1))
  End If

End Function
Public Function ArcCosJen(dblValue As Double) As Double

'  ArcCosJen = Atn(-dblValue / Sqr(-dblValue * dblValue + 1)) + 2 * Atn(1)
   
  If dblValue = 1 Then
    ArcCosJen = -dblPI / 2
  ElseIf dblValue = -1 Then
    ArcCosJen = dblPI / 2
  Else
    ArcCosJen = Atn(-dblValue / Sqr(-dblValue * dblValue + 1))
  End If
  
  ArcCosJen = ArcCosJen + (dblPI / 2)

End Function

Public Function DistanceVincentyPoints(pPoint1 As IPoint, pPoint2 As IPoint, dblAZ1 As Double, dblAZ2 As Double) As Double

  DistanceVincentyPoints = DistanceVincentyNumbers(pPoint1.x, pPoint1.Y, pPoint2.x, pPoint2.Y, dblAZ1, dblAZ2)

End Function

Public Function DistanceVincentyPoints2(pPoint1 As IPoint, pPoint2 As IPoint, dblAZ1 As Double, dblAZ2 As Double, _
  Optional dblEquatorialRadius As Double = 6378137, Optional dblPolarRadius As Double = 6356752.31424518) As Double
  
  ' MODIFICATION OF DistanceVincentyPoints, TO ALLOW FOR ANY ELLIPSOID
  
  DistanceVincentyPoints2 = DistanceVincentyNumbers2(pPoint1.x, pPoint1.Y, pPoint2.x, pPoint2.Y, dblAZ1, dblAZ2, _
        dblEquatorialRadius, dblPolarRadius)

End Function

Public Sub PointLineVincenty(pPoint As IPoint, dblLength As Double, dblAzimuth As Double, pNewPoint As IPoint, _
      dblAZ2 As Double, Optional lngNumVertices As Long, Optional pPolyline As IPolyline)
      
  Dim pWGS84 As IGeographicCoordinateSystem
  Dim pSpatRefFact As ISpatialReferenceFactory
  'Set the spatial reference factory to a new spatial reference environment
  Set pSpatRefFact = New SpatialReferenceEnvironment
  Set pWGS84 = pSpatRefFact.CreateGeographicCoordinateSystem(esriSRGeoCS_WGS1984)
  
  If lngNumVertices > 0 Then
    If lngNumVertices = 1 Then lngNumVertices = 2
    
    Dim dblShort As Double
    dblShort = dblLength / (lngNumVertices - 1)
    
    Dim pPointCollection As IPointCollection
    If pPolyline Is Nothing Then
      Set pPolyline = New Polyline
    End If
    Set pPointCollection = pPolyline
    
    ' ADD FIRST VERTEX
    pPointCollection.AddPoint pPoint
    
    ' ADD INTERNAL VERTICES
    If lngNumVertices > 2 Then
      Dim lngCounter As Long
      Dim dblCurrentDistance As Double
      Dim pTempPoint As IPoint
      For lngCounter = 1 To (lngNumVertices - 2)
        Set pTempPoint = New Point
        PointLineVincentyPerPoint pPoint, lngCounter * dblShort, dblAzimuth, pTempPoint, dblAZ2
        pPointCollection.AddPoint pTempPoint
      Next lngCounter
    End If
    
    ' ADD LAST VERTEX AND SET FINAL POINT AND AZIMUTH VALUES
    PointLineVincentyPerPoint pPoint, dblLength, dblAzimuth, pNewPoint, dblAZ2
    pPointCollection.AddPoint pNewPoint
    Set pPolyline.SpatialReference = pWGS84
    Set pNewPoint.SpatialReference = pWGS84
  Else
    PointLineVincentyPerPoint pPoint, dblLength, dblAzimuth, pNewPoint, dblAZ2
    Set pNewPoint.SpatialReference = pWGS84
  End If

End Sub

Public Sub PointLineVincenty2(pPoint As IPoint, dblLength As Double, dblAzimuth As Double, pNewPoint As IPoint, _
      dblAZ2 As Double, Optional lngNumVertices As Long, Optional pPolyline As IPolyline)
  
  ' MODIFICATION OF PointLineVincenty, TO ALLOW FOR ANY ELLIPSOID
  ' ASSUMES INCOMING POINT IS IN GEOGRAPHIC COORDINATE SYSTEM
  
  Dim pGCS As IGeographicCoordinateSystem
  Dim pSpRef As ISpatialReference
  
  Set pSpRef = pPoint.SpatialReference
  If TypeOf pSpRef Is IGeographicCoordinateSystem Then
    Set pGCS = pSpRef
  Else
    Dim pPrjCS As IProjectedCoordinateSystem
    Set pPrjCS = pSpRef
    Set pGCS = pPrjCS.GeographicCoordinateSystem
  End If
  
  Dim dblEquatorialRadius As Double
  Dim dblPolarRadius As Double
  Dim pEllipsoid As ISpheroid
  
  Set pEllipsoid = pGCS.Datum.Spheroid
  dblEquatorialRadius = pEllipsoid.SemiMajorAxis
  dblPolarRadius = pEllipsoid.SemiMinorAxis
  
  If lngNumVertices > 0 Then
    If lngNumVertices = 1 Then lngNumVertices = 2
    
    Dim dblShort As Double
    dblShort = dblLength / (lngNumVertices - 1)
    
    Dim pPointCollection As IPointCollection
    If pPolyline Is Nothing Then
      Set pPolyline = New Polyline
    End If
    Set pPointCollection = pPolyline
    
    ' ADD FIRST VERTEX
    pPointCollection.AddPoint pPoint
    
    ' ADD INTERNAL VERTICES
    If lngNumVertices > 2 Then
      Dim lngCounter As Long
      Dim dblCurrentDistance As Double
      Dim pTempPoint As IPoint
      For lngCounter = 1 To (lngNumVertices - 2)
        Set pTempPoint = New Point
        PointLineVincentyPerPoint2 pPoint, lngCounter * dblShort, dblAzimuth, pTempPoint, dblAZ2, dblEquatorialRadius, dblPolarRadius
        pPointCollection.AddPoint pTempPoint
      Next lngCounter
    End If
    
    ' ADD LAST VERTEX AND SET FINAL POINT AND AZIMUTH VALUES
    PointLineVincentyPerPoint2 pPoint, dblLength, dblAzimuth, pNewPoint, dblAZ2, dblEquatorialRadius, dblPolarRadius
    pPointCollection.AddPoint pNewPoint
    Set pPolyline.SpatialReference = pGCS
    Set pNewPoint.SpatialReference = pGCS
  Else
    PointLineVincentyPerPoint2 pPoint, dblLength, dblAzimuth, pNewPoint, dblAZ2, dblEquatorialRadius, dblPolarRadius
    Set pNewPoint.SpatialReference = pGCS
  End If

End Sub
Public Sub PointLineVincentyPerPoint2(pPoint As IPoint, dblLength As Double, dblAzimuth As Double, _
      pNewPoint As IPoint, dblAZ2 As Double, _
      Optional dblEquatorialRadius As Double = 6378137, Optional dblPolarRadius As Double = 6356752.31424518)
  
  ' MODIFICATION OF PointLineVincentyPerPoint, TO ALLOW FOR ANY ELLIPSOID
  
  ' ASSUMES pPoint IS GEOGRAPHIC

  ' ADAPTED FROM Vincenty, T. (1975). “Direct and inverse solutions of geodesics on the
  '                                    ellipsoid with application of nested equations.” Surv. Rev., XXII(176),
  '                                    88–93.
  ' ADAPTED FROM CHRIS VENESS; http://www.movable-type.co.uk/scripts/latlong-vincenty.html
  
  ' POINT 1 = dblPX, dblPY
  ' POINT 2 = dblQX, dblQY
  Dim dblPX As Double
  dblPX = pPoint.x
  Dim dblPY As Double
  dblPY = pPoint.Y
  
  If dblLength = 0 Then    ' SAME POINT
    pNewPoint.x = dblPX
    pNewPoint.Y = dblPY
    dblAZ2 = dblAzimuth
    Exit Sub
  End If
  
  ' WGS84 PARAMETERS ----------------------------------------
  Dim A As Double
  Dim B As Double
  Dim uSq As Double
  Dim dblA As Double
  Dim dblB As Double
  Dim f As Double
  Dim dblA1 As Double
  Dim dblB1 As Double
  
  Dim dblTanU1 As Double
  Dim dblSinU1 As Double
  Dim dblCosU1 As Double
  Dim U1 As Double          ' REDUCED LATITUDE FOR POINT 1;  dblPY
'  Dim U2 As Double          ' REDUCED LATITUDE FOR POINT 2;  dblQY
  
'  U2 = Atn((1 - f) * (Tan(DegToRad(dblQY))))
  
  A = dblEquatorialRadius   ' SPHEROID; EQUATORIAL RADIUS
  B = dblPolarRadius        ' SPHEROID; POLAR RADIUS
  f = (A - B) / A           ' FLATTENING
  
  dblTanU1 = (1 - f) * (Tan(DegToRad(dblPY)))
  U1 = Atn(dblTanU1)
  dblCosU1 = Cos(U1)
  dblSinU1 = Sin(U1)
  
  Dim s As Double
  s = dblLength
  
  Dim Sigma1 As Double
  Dim tanSigma1 As Double
  Dim cosAlpha1 As Double
  Dim sinAlpha1 As Double
'  Dim dblAlpha As Double                      ' AZIMUTH AT EQUATOR
  Dim cosAlpha As Double
  Dim cosSqAlpha As Double
  
  cosAlpha1 = Cos(DegToRad(dblAzimuth))
  sinAlpha1 = Sin(DegToRad(dblAzimuth))
  tanSigma1 = dblTanU1 / cosAlpha1                                                                    ' [1]
  Sigma1 = atan2(dblTanU1, cosAlpha1)
  Dim sinAlpha As Double
  sinAlpha = dblCosU1 * sinAlpha1                                                                  ' [2]
  cosSqAlpha = 1 - (sinAlpha ^ 2)                                                                  ' TRIG IDENTITY
  cosAlpha = Sqr(cosSqAlpha)
'  dblAlpha = ArcSinJen(sinAlpha)
  
  uSq = (cosSqAlpha * (A ^ 2 - B ^ 2)) / (B ^ 2)
  dblA1 = (uSq * (-768 + (uSq * (320 - (175 * uSq)))))
  dblA = 1 + ((uSq / 16384) * (4096 + dblA1))                                                      ' [3]
  dblB1 = (uSq * (-128 + (uSq * (74 - (uSq * 47)))))
  dblB = (uSq / 1024) * (256 + dblB1)                                                              ' [4]
  
  Dim Sigma As Double
  Dim sinSigma As Double
  Dim cosSigma As Double
  Dim DeltaSigma As Double
  Dim DeltaSigma1 As Double
  Dim DeltaSigma2 As Double
  Dim DeltaSigma3 As Double
  Dim cos2SigmaM As Double
  Dim C As Double
  Dim l As Double
  
  Dim lngIterations As Long
  lngIterations = 40
  
  Dim SigmaCompare As Double
  SigmaCompare = 2 * dblPI
  Sigma = s / (B * dblA)                  ' FIRST ESTIMATION
  
  Do While (Abs(Sigma - SigmaCompare) > 0.000000000001) And (lngIterations > 0)
    cos2SigmaM = Cos(2 * Sigma1 + Sigma)                                                             ' [5]
    sinSigma = Sin(Sigma)
    cosSigma = Cos(Sigma)
    DeltaSigma1 = ((dblB / 6) * cos2SigmaM * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2SigmaM ^ 2))
    DeltaSigma2 = ((dblB / 4) * (cosSigma * (-1 + 2 * cos2SigmaM ^ 2) - DeltaSigma1))
    DeltaSigma3 = cos2SigmaM + DeltaSigma2
    DeltaSigma = dblB * sinSigma * DeltaSigma3                                                       ' [6]
    SigmaCompare = Sigma
    Sigma = (s / (B * dblA)) + DeltaSigma                                                            ' [7]
    
    If lngIterations = 0 Then
      MsgBox "Vincenty Formula failed to converge!"
      Exit Sub
    End If
    lngIterations = lngIterations - 1
  Loop
  
  cos2SigmaM = Cos(2 * Sigma1 + Sigma)
  sinSigma = Sin(Sigma)
  cosSigma = Cos(Sigma)
  Dim dblLat2Denom As Double
  Dim dblLat2Temp As Double
  dblLat2Temp = dblSinU1 * sinSigma - dblCosU1 * cosSigma * cosAlpha1
  dblLat2Denom = (1 - f) * (Sqr(sinAlpha ^ 2 + dblLat2Temp ^ 2))
  
  ' CALCULATE LATITUDE FOR NEW POINT
  Dim dblLat2 As Double
  dblLat2 = atan2(dblSinU1 * cosSigma + dblCosU1 * sinSigma * cosAlpha1, dblLat2Denom)                ' [8]
  
  ' CALCULATE LONGITUDE FOR NEW POINT
  Dim dblLambda As Double
  Dim dblLambda1 As Double
  Dim dblLambda1a As Double
  dblLambda = atan2(sinSigma * sinAlpha1, dblCosU1 * cosSigma - dblSinU1 * sinSigma * cosAlpha1)      ' [9]
  C = (f / 16) * cosSqAlpha * (4 + (f * (4 - (3 * cosSqAlpha))))                                      ' [10]
  dblLambda1 = cos2SigmaM + C * cosSigma * (-1# + 2# * cos2SigmaM ^ 2#)
  dblLambda1a = C * sinSigma * dblLambda1
  Dim dblLambda2 As Double
  dblLambda2 = Sigma + dblLambda1a
  l = dblLambda - ((1 - C) * f * sinAlpha * dblLambda2)                                               ' [11]
  
  pNewPoint.x = dblPX + RadToDeg(l)
  pNewPoint.Y = RadToDeg(dblLat2)
  
  dblAZ2 = RadToDeg(atan2(sinAlpha, -dblLat2Temp))
  If dblAZ2 < 0 Then dblAZ2 = 360 + dblAZ2

End Sub

Public Sub PointLineVincentyPerPoint(pPoint As IPoint, dblLength As Double, dblAzimuth As Double, _
      pNewPoint As IPoint, dblAZ2 As Double)
      
  ' ASSUMES pPoint IS GEOGRAPHIC

  ' ADAPTED FROM Vincenty, T. (1975). “Direct and inverse solutions of geodesics on the
  '                                    ellipsoid with application of nested equations.” Surv. Rev., XXII(176),
  '                                    88–93.
  ' ADAPTED FROM CHRIS VENESS; http://www.movable-type.co.uk/scripts/latlong-vincenty.html
  
  ' POINT 1 = dblPX, dblPY
  ' POINT 2 = dblQX, dblQY
  Dim dblPX As Double
  dblPX = pPoint.x
  Dim dblPY As Double
  dblPY = pPoint.Y
  
  If dblLength = 0 Then    ' SAME POINT
    pNewPoint.x = dblPX
    pNewPoint.Y = dblPY
    dblAZ2 = dblAzimuth
    Exit Sub
  End If
  
  ' WGS84 PARAMETERS ----------------------------------------
  Dim A As Double
  Dim B As Double
  Dim uSq As Double
  Dim dblA As Double
  Dim dblB As Double
  Dim f As Double
  Dim dblA1 As Double
  Dim dblB1 As Double
  
  Dim dblTanU1 As Double
  Dim dblSinU1 As Double
  Dim dblCosU1 As Double
  Dim U1 As Double          ' REDUCED LATITUDE FOR POINT 1;  dblPY
'  Dim U2 As Double          ' REDUCED LATITUDE FOR POINT 2;  dblQY
  
'  U2 = Atn((1 - f) * (Tan(DegToRad(dblQY))))
  
  A = 6378137               ' WGS84 SPHEROID; EQUATORIAL RADIUS
  B = 6356752.31424518      ' WGS84 SPHEROID; POLAR RADIUS
  f = 1 / 298.257223563     ' WGS84 SPHEROID; FLATTENING
  
  dblTanU1 = (1 - f) * (Tan(DegToRad(dblPY)))
  U1 = Atn(dblTanU1)
  dblCosU1 = Cos(U1)
  dblSinU1 = Sin(U1)
  
  Dim s As Double
  s = dblLength
  
  Dim Sigma1 As Double
  Dim tanSigma1 As Double
  Dim cosAlpha1 As Double
  Dim sinAlpha1 As Double
'  Dim dblAlpha As Double                      ' AZIMUTH AT EQUATOR
  Dim cosAlpha As Double
  Dim cosSqAlpha As Double
  
  cosAlpha1 = Cos(DegToRad(dblAzimuth))
  sinAlpha1 = Sin(DegToRad(dblAzimuth))
  tanSigma1 = dblTanU1 / cosAlpha1                                                                    ' [1]
  Sigma1 = atan2(dblTanU1, cosAlpha1)
  Dim sinAlpha As Double
  sinAlpha = dblCosU1 * sinAlpha1                                                                  ' [2]
  cosSqAlpha = 1 - (sinAlpha ^ 2)                                                                  ' TRIG IDENTITY
  cosAlpha = Sqr(cosSqAlpha)
'  dblAlpha = ArcSinJen(sinAlpha)
  
  uSq = (cosSqAlpha * (A ^ 2 - B ^ 2)) / (B ^ 2)
  dblA1 = (uSq * (-768 + (uSq * (320 - (175 * uSq)))))
  dblA = 1 + ((uSq / 16384) * (4096 + dblA1))                                                      ' [3]
  dblB1 = (uSq * (-128 + (uSq * (74 - (uSq * 47)))))
  dblB = (uSq / 1024) * (256 + dblB1)                                                              ' [4]
  
  Dim Sigma As Double
  Dim sinSigma As Double
  Dim cosSigma As Double
  Dim DeltaSigma As Double
  Dim DeltaSigma1 As Double
  Dim DeltaSigma2 As Double
  Dim DeltaSigma3 As Double
  Dim cos2SigmaM As Double
  Dim C As Double
  Dim l As Double
  
  Dim lngIterations As Long
  lngIterations = 40
  
  Dim SigmaCompare As Double
  SigmaCompare = 2 * dblPI
  Sigma = s / (B * dblA)                  ' FIRST ESTIMATION
  
  Do While (Abs(Sigma - SigmaCompare) > 0.000000000001) And (lngIterations > 0)
    cos2SigmaM = Cos(2 * Sigma1 + Sigma)                                                             ' [5]
    sinSigma = Sin(Sigma)
    cosSigma = Cos(Sigma)
    DeltaSigma1 = ((dblB / 6) * cos2SigmaM * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2SigmaM ^ 2))
    DeltaSigma2 = ((dblB / 4) * (cosSigma * (-1 + 2 * cos2SigmaM ^ 2) - DeltaSigma1))
    DeltaSigma3 = cos2SigmaM + DeltaSigma2
    DeltaSigma = dblB * sinSigma * DeltaSigma3                                                       ' [6]
    SigmaCompare = Sigma
    Sigma = (s / (B * dblA)) + DeltaSigma                                                            ' [7]
    
    If lngIterations = 0 Then
      MsgBox "Vincenty Formula failed to converge!"
      Exit Sub
    End If
    lngIterations = lngIterations - 1
  Loop
  
  cos2SigmaM = Cos(2 * Sigma1 + Sigma)
  sinSigma = Sin(Sigma)
  cosSigma = Cos(Sigma)
  Dim dblLat2Denom As Double
  Dim dblLat2Temp As Double
  dblLat2Temp = dblSinU1 * sinSigma - dblCosU1 * cosSigma * cosAlpha1
  dblLat2Denom = (1 - f) * (Sqr(sinAlpha ^ 2 + dblLat2Temp ^ 2))
  
  ' CALCULATE LATITUDE FOR NEW POINT
  Dim dblLat2 As Double
  dblLat2 = atan2(dblSinU1 * cosSigma + dblCosU1 * sinSigma * cosAlpha1, dblLat2Denom)                ' [8]
  
  ' CALCULATE LONGITUDE FOR NEW POINT
  Dim dblLambda As Double
  Dim dblLambda1 As Double
  Dim dblLambda1a As Double
  dblLambda = atan2(sinSigma * sinAlpha1, dblCosU1 * cosSigma - dblSinU1 * sinSigma * cosAlpha1)      ' [9]
  C = (f / 16) * cosSqAlpha * (4 + (f * (4 - (3 * cosSqAlpha))))                                      ' [10]
  dblLambda1 = cos2SigmaM + C * cosSigma * (-1# + 2# * cos2SigmaM ^ 2#)
  dblLambda1a = C * sinSigma * dblLambda1
  Dim dblLambda2 As Double
  dblLambda2 = Sigma + dblLambda1a
  l = dblLambda - ((1 - C) * f * sinAlpha * dblLambda2)                                               ' [11]
  
  pNewPoint.x = dblPX + RadToDeg(l)
  pNewPoint.Y = RadToDeg(dblLat2)
  
  dblAZ2 = RadToDeg(atan2(sinAlpha, -dblLat2Temp))
  If dblAZ2 < 0 Then dblAZ2 = 360 + dblAZ2

End Sub

Public Function DistanceVincentyNumbers(dblPX As Double, dblPY As Double, dblQX As Double, dblQY As Double, _
        dblAZ1 As Double, dblAZ2 As Double) As Double
  
  ' ADAPTED FROM Vincenty, T. (1975). “Direct and inverse solutions of geodesics on the
  '                                    ellipsoid with application of nested equations.” Surv. Rev., XXII(176),
  '                                    88–93.
  ' ADAPTED FROM CHRIS VENESS; http://www.movable-type.co.uk/scripts/latlong-vincenty-direct.html
  
  ' POINT 1 = dblPX, dblPY
  ' POINT 2 = dblQX, dblQY
  
  If dblPX = dblQX And dblPY = dblQY Then      ' SAME POINT
    DistanceVincentyNumbers = 0
    dblAZ1 = 0
    dblAZ2 = 0
    Exit Function
  End If
  
  
  Dim A As Double
  Dim B As Double
  A = 6378137               ' WGS84 SPHEROID; EQUATORIAL RADIUS
  B = 6356752.31424518      ' WGS84 SPHEROID; POLAR RADIUS
  
  Dim f As Double
  f = 1 / 298.257223563     ' WGS84 SPHEROID; FLATTENING
  
  Dim l As Double           ' DIFFERENCE IN LONGITUDE
  l = DegToRad(dblQX - dblPX)
  
  Dim U1 As Double          ' REDUCED LATITUDE FOR POINT 1;  dblPY
  Dim U2 As Double          ' REDUCED LATITUDE FOR POINT 2;  dblQY
  
  U1 = Atn((1 - f) * (Tan(DegToRad(dblPY))))
  U2 = Atn((1 - f) * (Tan(DegToRad(dblQY))))
  
  Dim dblSinU1 As Double
  Dim dblSinU2 As Double
  Dim dblCosU1 As Double
  Dim dblCosU2 As Double
  
  dblSinU1 = Sin(U1)
  dblCosU1 = Cos(U1)
  dblSinU2 = Sin(U2)
  dblCosU2 = Cos(U2)
  
  Dim dblLambda As Double, dblLambdaComp As Double
  dblLambda = l                     ' FIRST ESTIMATE OF LAMBDA
  dblLambdaComp = 2 * dblPI
  Dim lngIterations As Long
  lngIterations = 40
  
  Dim sinLambda As Double
  Dim cosLambda As Double
  Dim sinSigma As Double
  Dim cosSigma As Double
  Dim Sigma As Double
  Dim sinAlpha As Double
  Dim cosSqAlpha As Double
  Dim cos2SigmaM As Double
  Dim C As Double
  
  Dim dblLambda1 As Double
  Dim dblLambda1a As Double
  
  Do While (Abs(dblLambda - dblLambdaComp) > 0.000000000001) And (lngIterations > 0)       ' VINCENTY EQUATION NUMBERS
    sinLambda = Sin(dblLambda)                                                          '  |
    cosLambda = Cos(dblLambda)                                                          '  |
    sinSigma = Sqr((dblCosU2 * sinLambda) ^ 2 + ((dblCosU1 * dblSinU2) - _
            (dblSinU1 * dblCosU2 * cosLambda)) ^ 2)                                     ' [14]
    cosSigma = (dblSinU1 * dblSinU2) + (dblCosU1 * dblCosU2 * cosLambda)                ' [15]
    Sigma = atan2(sinSigma, cosSigma)                                                   ' [16]
    sinAlpha = (dblCosU1 * dblCosU2 * sinLambda) / sinSigma                             ' [17]
    cosSqAlpha = 1 - (sinAlpha ^ 2)                                                     ' TRIG IDENTITY
    If cosSqAlpha = 0 Then
      cos2SigmaM = 0                                                                    ' ADAPTED FROM VENESS
    Else
      cos2SigmaM = cosSigma - ((2 * dblSinU1 * dblSinU2) / cosSqAlpha)                  ' [18]
    End If
    
    C = (f / 16) * cosSqAlpha * (4 + (f * (4 - (3 * cosSqAlpha))))                      ' [10]
    dblLambdaComp = dblLambda
    dblLambda1 = cos2SigmaM + C * cosSigma * (-1 + (2 * cos2SigmaM * cos2SigmaM))
    dblLambda1a = C * sinSigma * dblLambda1
    ' VINCENTY WRITES EQUATION AS "L = dblLambda - ((1 - C)...
    dblLambda = l + ((1 - C) * f * sinAlpha * (Sigma + dblLambda1a))                    ' [11]
    
    If lngIterations = 0 Then
      MsgBox "Vincenty Formula failed to converge!"
      DistanceVincentyNumbers = -9999
      Exit Function
    End If
    lngIterations = lngIterations - 1
  Loop
  
  Dim uSq As Double
  Dim dblA As Double
  Dim dblB As Double
  Dim DeltaSigma As Double
  Dim s As Double
  
  Dim DeltaSigma1 As Double
  Dim DeltaSigma2 As Double
  Dim DeltaSigma3 As Double
  
  uSq = (cosSqAlpha * (A ^ 2 - B ^ 2)) / (B ^ 2)
  
  Dim dblA1 As Double
  Dim dblB1 As Double
  
  dblA1 = (uSq * (-768 + (uSq * (320 - (175 * uSq)))))
  dblA = 1 + ((uSq / 16384) * (4096 + dblA1))                                           ' [3]
  dblB1 = (uSq * (-128 + (uSq * (74 - (uSq * 47)))))
  dblB = (uSq / 1024) * (256 + dblB1)              ' [4]
  
  DeltaSigma1 = ((dblB / 6) * cos2SigmaM * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2SigmaM ^ 2))
  DeltaSigma2 = ((dblB / 4) * (cosSigma * (-1 + 2 * cos2SigmaM ^ 2) - DeltaSigma1))
  DeltaSigma3 = cos2SigmaM + DeltaSigma2
  DeltaSigma = dblB * sinSigma * DeltaSigma3                                                                 ' [6]
  s = B * dblA * (Sigma - DeltaSigma)
  
'  var uSq = cosSqAlpha * (a*a - b*b) / (b*b);
'  var A = 1 + uSq/16384*(4096+uSq*(-768+uSq*(320-175*uSq)));
'  var B = uSq/1024 * (256+uSq*(-128+uSq*(74-47*uSq)));
'  var deltaSigma = B*sinSigma*(cos2SigmaM+B/4*(cosSigma*(-1+2*cos2SigmaM*cos2SigmaM)-
'    B/6*cos2SigmaM*(-3+4*sinSigma*sinSigma)*(-3+4*cos2SigmaM*cos2SigmaM)));
'  var s = b*A*(sigma-deltaSigma);
'
'  s = s.toFixed(3); // round to 1mm precision
'  return s;
'}
  
  
  dblAZ1 = RadToDeg(atan2(dblCosU2 * sinLambda, (dblCosU1 * dblSinU2) - (dblSinU1 * dblCosU2 * cosLambda)))
  dblAZ2 = RadToDeg(atan2(dblCosU1 * sinLambda, -(dblSinU1 * dblCosU2) + (dblCosU1 * dblSinU2 * cosLambda)))
  
  If dblAZ1 < 0 Then dblAZ1 = 360 + dblAZ1
  If dblAZ2 < 0 Then dblAZ2 = 360 + dblAZ2
  DistanceVincentyNumbers = s

End Function

Public Function DistanceVincentyNumbers2(dblPX As Double, dblPY As Double, dblQX As Double, dblQY As Double, _
        dblAZ1 As Double, dblAZ2 As Double, _
        Optional dblEquatorialRadius As Double = 6378137, Optional dblPolarRadius As Double = 6356752.31424518) As Double
  
  ' MODIFICATION OF DistanceVincentyNumbers TO ALLOW FOR ANY ELLIPSOID
  
  ' ADAPTED FROM Vincenty, T. (1975). “Direct and inverse solutions of geodesics on the
  '                                    ellipsoid with application of nested equations.” Surv. Rev., XXII(176),
  '                                    88–93.
  ' ADAPTED FROM CHRIS VENESS; http://www.movable-type.co.uk/scripts/latlong-vincenty-direct.html
  
  ' POINT 1 = dblPX, dblPY
  ' POINT 2 = dblQX, dblQY
  
  If dblPX = dblQX And dblPY = dblQY Then      ' SAME POINT
    DistanceVincentyNumbers2 = 0
    dblAZ1 = 0
    dblAZ2 = 0
    Exit Function
  End If
  
  
  Dim A As Double
  Dim B As Double
  A = dblEquatorialRadius   ' SPHEROID; EQUATORIAL RADIUS
  B = dblPolarRadius        ' SPHEROID; POLAR RADIUS
  
  Dim f As Double
  f = (A - B) / A           ' FLATTENING
  
  Dim l As Double           ' DIFFERENCE IN LONGITUDE
  l = DegToRad(dblQX - dblPX)
  
  Dim U1 As Double          ' REDUCED LATITUDE FOR POINT 1;  dblPY
  Dim U2 As Double          ' REDUCED LATITUDE FOR POINT 2;  dblQY
  
  U1 = Atn((1 - f) * (Tan(DegToRad(dblPY))))
  U2 = Atn((1 - f) * (Tan(DegToRad(dblQY))))
  
  Dim dblSinU1 As Double
  Dim dblSinU2 As Double
  Dim dblCosU1 As Double
  Dim dblCosU2 As Double
  
  dblSinU1 = Sin(U1)
  dblCosU1 = Cos(U1)
  dblSinU2 = Sin(U2)
  dblCosU2 = Cos(U2)
  
  Dim dblLambda As Double, dblLambdaComp As Double
  dblLambda = l                     ' FIRST ESTIMATE OF LAMBDA
  dblLambdaComp = 2 * dblPI
  Dim lngIterations As Long
  lngIterations = 40
  
  Dim sinLambda As Double
  Dim cosLambda As Double
  Dim sinSigma As Double
  Dim cosSigma As Double
  Dim Sigma As Double
  Dim sinAlpha As Double
  Dim cosSqAlpha As Double
  Dim cos2SigmaM As Double
  Dim C As Double
  
  Dim dblLambda1 As Double
  Dim dblLambda1a As Double
  
  Do While (Abs(dblLambda - dblLambdaComp) > 0.000000000001) And (lngIterations > 0)       ' VINCENTY EQUATION NUMBERS
    sinLambda = Sin(dblLambda)                                                          '  |
    cosLambda = Cos(dblLambda)                                                          '  |
    sinSigma = Sqr((dblCosU2 * sinLambda) ^ 2 + ((dblCosU1 * dblSinU2) - _
            (dblSinU1 * dblCosU2 * cosLambda)) ^ 2)                                     ' [14]
    cosSigma = (dblSinU1 * dblSinU2) + (dblCosU1 * dblCosU2 * cosLambda)                ' [15]
    Sigma = atan2(sinSigma, cosSigma)                                                   ' [16]
    sinAlpha = (dblCosU1 * dblCosU2 * sinLambda) / sinSigma                             ' [17]
    cosSqAlpha = 1 - (sinAlpha ^ 2)                                                     ' TRIG IDENTITY
    If cosSqAlpha = 0 Then
      cos2SigmaM = 0                                                                    ' ADAPTED FROM VENESS
    Else
      cos2SigmaM = cosSigma - ((2 * dblSinU1 * dblSinU2) / cosSqAlpha)                  ' [18]
    End If
    
    C = (f / 16) * cosSqAlpha * (4 + (f * (4 - (3 * cosSqAlpha))))                      ' [10]
    dblLambdaComp = dblLambda
    dblLambda1 = cos2SigmaM + C * cosSigma * (-1 + (2 * cos2SigmaM * cos2SigmaM))
    dblLambda1a = C * sinSigma * dblLambda1
    ' VINCENTY WRITES EQUATION AS "L = dblLambda - ((1 - C)...
    dblLambda = l + ((1 - C) * f * sinAlpha * (Sigma + dblLambda1a))                    ' [11]
    
    If lngIterations = 0 Then
      MsgBox "Vincenty Formula failed to converge!"
      DistanceVincentyNumbers2 = -9999
      Exit Function
    End If
    lngIterations = lngIterations - 1
  Loop
  
  Dim uSq As Double
  Dim dblA As Double
  Dim dblB As Double
  Dim DeltaSigma As Double
  Dim s As Double
  
  Dim DeltaSigma1 As Double
  Dim DeltaSigma2 As Double
  Dim DeltaSigma3 As Double
  
  uSq = (cosSqAlpha * (A ^ 2 - B ^ 2)) / (B ^ 2)
  
  Dim dblA1 As Double
  Dim dblB1 As Double
  
  dblA1 = (uSq * (-768 + (uSq * (320 - (175 * uSq)))))
  dblA = 1 + ((uSq / 16384) * (4096 + dblA1))                                           ' [3]
  dblB1 = (uSq * (-128 + (uSq * (74 - (uSq * 47)))))
  dblB = (uSq / 1024) * (256 + dblB1)              ' [4]
  
  DeltaSigma1 = ((dblB / 6) * cos2SigmaM * (-3 + 4 * sinSigma ^ 2) * (-3 + 4 * cos2SigmaM ^ 2))
  DeltaSigma2 = ((dblB / 4) * (cosSigma * (-1 + 2 * cos2SigmaM ^ 2) - DeltaSigma1))
  DeltaSigma3 = cos2SigmaM + DeltaSigma2
  DeltaSigma = dblB * sinSigma * DeltaSigma3                                                                 ' [6]
  s = B * dblA * (Sigma - DeltaSigma)
  
'  var uSq = cosSqAlpha * (a*a - b*b) / (b*b);
'  var A = 1 + uSq/16384*(4096+uSq*(-768+uSq*(320-175*uSq)));
'  var B = uSq/1024 * (256+uSq*(-128+uSq*(74-47*uSq)));
'  var deltaSigma = B*sinSigma*(cos2SigmaM+B/4*(cosSigma*(-1+2*cos2SigmaM*cos2SigmaM)-
'    B/6*cos2SigmaM*(-3+4*sinSigma*sinSigma)*(-3+4*cos2SigmaM*cos2SigmaM)));
'  var s = b*A*(sigma-deltaSigma);
'
'  s = s.toFixed(3); // round to 1mm precision
'  return s;
'}
  
  
  dblAZ1 = RadToDeg(atan2(dblCosU2 * sinLambda, (dblCosU1 * dblSinU2) - (dblSinU1 * dblCosU2 * cosLambda)))
  dblAZ2 = RadToDeg(atan2(dblCosU1 * sinLambda, -(dblSinU1 * dblCosU2) + (dblCosU1 * dblSinU2 * cosLambda)))
  
  If dblAZ1 < 0 Then dblAZ1 = 360 + dblAZ1
  If dblAZ2 < 0 Then dblAZ2 = 360 + dblAZ2
  DistanceVincentyNumbers2 = s

End Function

Public Function RadToDeg(dblRad As Double) As Double

  RadToDeg = dblRad * 180 / dblPI

End Function
Public Function DegToRad(dblDeg As Double) As Double

  DegToRad = dblDeg * dblPI / 180

End Function

Public Function atan2(dblDeltaY As Double, dblDeltaX As Double) As Double

'  If X > 0 Then
'    atan2 = Atn(Y / X)
'  ElseIf X < 0 Then
'    If Y = 0 Then
'      atan2 = (dblPi - Atn(Abs(Y / X)))
'    Else
'      atan2 = Sgn(Y) * (dblPi - Atn(Abs(Y / X)))
'    End If
'  Else    ' IF X = 0
'    If Y = 0 Then
'      atan2 = 0
'    Else
'      atan2 = Sgn(Y) * dblPi / 2
'    End If
'  End If

  
  If dblDeltaX > 0 Then
    atan2 = Atn(dblDeltaY / dblDeltaX)
  ElseIf dblDeltaX < 0 Then
    If dblDeltaY = 0 Then
      atan2 = dblPI
    Else
      atan2 = Sgn(dblDeltaY) * (dblPI - Atn(Abs(dblDeltaY / dblDeltaX)))
    End If
  Else    ' IF dblDeltaX  = 0
    If dblDeltaY = 0 Then
      atan2 = 0
    Else
      atan2 = Sgn(dblDeltaY) * dblPI / 2
    End If
  End If
  
End Function
Public Function TriangleAreaLegs(dblLeg1 As Double, dblLeg2 As Double, dblLeg3 As Double) As Double

  Dim dblS As Double
  dblS = (dblLeg1 + dblLeg2 + dblLeg3) / 2
  TriangleAreaLegs = Sqr(dblS * (dblS - dblLeg1) * (dblS - dblLeg2) * (dblS - dblLeg3))

End Function
Public Function TriangleAreaPoints(pPoint1 As IPoint, pPoint2 As IPoint, pPoint3 As IPoint) As Double

  TriangleAreaPoints = Abs(((((pPoint2.x - pPoint1.x) * (pPoint3.Y - pPoint1.Y)) - ((pPoint3.x - pPoint1.x) * (pPoint2.Y - pPoint1.Y))) / 2))

End Function

Public Function TriangleAreaPointsValues(dbl1X As Double, dbl1Y As Double, _
                                         dbl2X As Double, dbl2Y As Double, _
                                         dbl3X As Double, dbl3Y As Double) As Double

  TriangleAreaPointsValues = Abs(((((dbl2X - dbl1X) * (dbl3Y - dbl1Y)) - ((dbl3X - dbl1X) * (dbl2Y - dbl1Y))) / 2))

End Function

Public Function TriangleAreaPoints3D(pPoint1 As IPoint, pPoint2 As IPoint, pPoint3 As IPoint) As Double
  
  TriangleAreaPoints3D = TriangleAreaPoints3DValues(pPoint1.x, pPoint1.Y, pPoint1.Z, pPoint2.x, pPoint2.Y, pPoint2.Z, _
                                                  pPoint3.x, pPoint3.Y, pPoint3.Z)

End Function



Public Function TriangleAreaPoints3DValues(dblPX As Double, dblPY As Double, dblPZ As Double, _
                                           dblQX As Double, dblQY As Double, dblQZ As Double, _
                                           dblRX As Double, dblRY As Double, dblRZ As Double) As Double

  Dim dblI As Double
  Dim dblJ As Double
  Dim dblK As Double
  
  dblI = (((dblQY - dblPY) * (dblRZ - dblPZ)) - ((dblRY - dblPY) * (dblQZ - dblPZ))) ^ 2
  dblJ = (((dblQX - dblPX) * (dblRZ - dblPZ)) - ((dblRX - dblPX) * (dblQZ - dblPZ))) ^ 2
  dblK = (((dblQX - dblPX) * (dblRY - dblPY)) - ((dblRX - dblPX) * (dblQY - dblPY))) ^ 2

  TriangleAreaPoints3DValues = (Sqr(dblI + dblJ + dblK)) / 2

End Function
                                           

Public Function EnvelopeToPolygon(pEnv As IEnvelope) As IPolygon

  Dim pPtColl As IPointCollection

  Set pPtColl = New Polygon
  With pPtColl
      .AddPoint pEnv.LowerLeft
      .AddPoint pEnv.UpperLeft
      .AddPoint pEnv.UpperRight
      .AddPoint pEnv.LowerRight
      'Close the polygon
      .AddPoint pEnv.LowerLeft
  End With
  
  Dim pPolygon As IPolygon
  Set pPolygon = pPtColl
  Set pPolygon.SpatialReference = pEnv.SpatialReference
  Dim pTopoOp As ITopologicalOperator
  Set pTopoOp = pPolygon
  pTopoOp.Simplify
    
  Set EnvelopeToPolygon = pPtColl

End Function


Public Function EllipticArcToPolygon2(SegCollection As ISegmentCollection, NumVertices As Long) As IMultipoint
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
  
'  Dim pEllArc As IEllipticArc
      
On Error GoTo erh

  Dim pCurve As ICurve
  Dim pGeometry As IGeometry
  
  Dim anIndex As Long
  Dim lngSegCount As Long
  lngSegCount = SegCollection.SegmentCount - 1
  Dim theLength As Double
  theLength = 0
  Dim theTestLength As Double
  Dim lngLengths() As Long
  ReDim lngLengths(lngSegCount)
  For anIndex = 0 To lngSegCount
    theTestLength = SegCollection.Segment(anIndex).length
    theLength = theLength + theTestLength
    lngLengths(anIndex) = theTestLength
  Next anIndex
  
  Dim pProportion As Double
  Dim lngVertices() As Long
  Dim lngNumVertices As Long
  ReDim lngVertices(lngSegCount)
  For anIndex = 0 To lngSegCount
    lngNumVertices = Int((lngLengths(anIndex) / theLength) * NumVertices)
    If lngNumVertices < 8 Then lngNumVertices = 8
    lngVertices(anIndex) = lngNumVertices
  Next anIndex
  
  Dim pMpt As IPointCollection
  Set pMpt = New Multipoint
  Dim pPoint As IPoint
  Set pPoint = New Point
  Dim pClone As IClone
  
  Dim pRatio As Double
  Dim anIndex2 As Long
  
  For anIndex = 0 To lngSegCount
    lngNumVertices = lngVertices(anIndex)
    pRatio = 1 / lngNumVertices
    Set pCurve = SegCollection.Segment(anIndex)
    
    For anIndex2 = 0 To lngNumVertices
'      If pGeometry.GeometryType = esriGeometryEllipticArc Then
      pCurve.QueryPoint 0, (pRatio * anIndex2), True, pPoint
      Set pClone = pPoint
        
 '   Graphic_MakeFromGeometry pMxDoc, pPoint, "DeleteMe"
    
      pMpt.AddPoint pClone.Clone
    Next anIndex2
  Next anIndex
  
  Set EllipticArcToPolygon2 = pMpt
    Exit Function
  
erh:
    MsgBox "Failed in EllipticArcToPolygon2: " & err.Description
End Function

Public Function CurveToPolygon(pOrigGeometry As IGeometry, NumVertices As Long) As IPolygon
On Error GoTo erh
  
  Dim pGeometryCollection As IGeometryCollection
  Set pGeometryCollection = pOrigGeometry
  Dim pSpRef As ISpatialReference
  Set pSpRef = pOrigGeometry.SpatialReference
  
  Dim pOrigPolygon As IPolycurve
  Set pOrigPolygon = pOrigGeometry
  
  Dim dblFullLength As Double
  dblFullLength = pOrigPolygon.length
  
  Dim pCurve As ICurve
  Dim pGeometry As IGeometry
  Dim pPolygon As IPointCollection
  Dim pRing As IRing
  Dim pSegment As ISegment
  Dim pStartPoint As IPoint
  Set pStartPoint = New Point
  Dim pEndPoint As IPoint
  Set pEndPoint = New Point
  Dim pClone As IClone
  Dim booFoundCurve As Boolean
  
  Dim lngRingCount As Long
  Dim lngNumVertices As Long
  Dim pRatio As Double
  Dim anIndex As Long
  Dim anIndex2 As Long
  Dim anIndex3 As Long
  Dim lngSegCount As Long
  
  Dim SegCollection As ISegmentCollection
  Dim pNewSegCollection As ISegmentCollection
  
  Dim pPathSegColl As ISegmentCollection
  Dim pNewSegment As ISegment
  Dim pNewLine As esriGeometry.ILine
  
  Dim pNewPolyGeoColl As IGeometryCollection
  Set pNewPolyGeoColl = New Polygon
  Dim pNewRingGeometry As IGeometry
  Dim pPath As IPath
  Dim pSegmentCollection As ISegmentCollection
  Dim pNewSegCol As ISegmentCollection
  
  lngRingCount = pGeometryCollection.GeometryCount - 1
  For anIndex = 0 To lngRingCount
    If TypeOf pOrigGeometry Is IPolyline Then
      Set pPath = pGeometryCollection.Geometry(anIndex)
      Set pSegmentCollection = pPath
      Set pNewSegCol = New Ring
      pNewSegCol.AddSegmentCollection pSegmentCollection
      Set pRing = pNewSegCol
      pRing.Close
    Else
      Set pRing = pGeometryCollection.Geometry(anIndex)
    End If
    Set SegCollection = pRing
    Set pNewSegCollection = New Ring
    lngSegCount = SegCollection.SegmentCount - 1
    For anIndex2 = 0 To lngSegCount
      Set pSegment = SegCollection.Segment(anIndex2)
      Set pGeometry = pSegment
      If pGeometry.GeometryType <> esriGeometryLine Then ' IF SEGMENT IS CURVE
        booFoundCurve = True
        lngNumVertices = Int((pSegment.length / dblFullLength) * NumVertices)
        If lngNumVertices < 8 Then lngNumVertices = 8
        pRatio = 1 / lngNumVertices
       
        Set pCurve = pSegment
        Set pPathSegColl = New Path
        Set pNewSegment = New esriGeometry.Line
        Set pStartPoint = pCurve.FromPoint
        For anIndex3 = 1 To lngNumVertices
          pCurve.QueryPoint 0, (pRatio * anIndex3), True, pEndPoint
          pNewSegment.FromPoint = pStartPoint
          pNewSegment.ToPoint = pEndPoint
          
          Set pClone = pNewSegment
          pPathSegColl.AddSegment pClone.Clone
          
'          If anIndex3 < 4 Then
'            MsgBox "Start Point:  X = " & CStr(pStartPoint.X) & ", Y = " & CStr(pStartPoint.Y) & vbCrLf & _
'              "End Point:  X = " & CStr(pEndPoint.X) & ", Y = " & CStr(pEndPoint.Y) & vbCrLf & _
'              "Segment Length = " & CStr(pNewSegment.length) & vbCrLf & _
'              "Segment Collection Count = " & CStr(pPathSegColl.SegmentCount)
'          End If
          
          Set pClone = pEndPoint
          Set pStartPoint = pClone.Clone
          
        Next anIndex3
        pNewSegCollection.AddSegmentCollection pPathSegColl

      Else      ' IF SEGMENT IS ACTUALLY LINE, DON'T ADD MIDPOINTS
        Set pClone = pSegment
        pNewSegCollection.AddSegment pClone.Clone
      End If
    Next anIndex2
    Set pNewRingGeometry = pNewSegCollection
    pNewPolyGeoColl.AddGeometry pNewRingGeometry

  Next anIndex
  
  Dim pNewPolygon As IPolygon
  
  If booFoundCurve Or (TypeOf pOrigGeometry Is IPolyline) Then
  
    Set pNewPolygon = pNewPolyGeoColl
    Dim pTopoOp As ITopologicalOperator
    Set pTopoOp = pNewPolygon
    pTopoOp.Simplify
    Set pNewPolygon.SpatialReference = pSpRef

  Else
    Set pNewPolygon = pOrigGeometry
    Set pNewPolygon.SpatialReference = pSpRef
  End If
    
  Set CurveToPolygon = pNewPolygon
  Exit Function
  
erh:
    MsgBox "Failed in CurveToPolygon: " & vbCrLf & "Error = " & err.Description & vbCrLf & "Line Number = " & CStr(Erl)
End Function
Public Function CurveToPolyline(pOrigGeometry As IGeometry, NumVertices As Long) As IPolyline
On Error GoTo erh

  Dim pGeometryCollection As IGeometryCollection
  Set pGeometryCollection = pOrigGeometry
  Dim pSpRef As ISpatialReference
  Set pSpRef = pOrigGeometry.SpatialReference
  
  Dim pOrigPolyline As IPolycurve
  
  Set pOrigPolyline = pOrigGeometry
  
  Dim dblFullLength As Double
  dblFullLength = pOrigPolyline.length

  Dim pPath As IPath
  
  Dim pCurve As ICurve
  Dim pGeometry As IGeometry
  Dim pSegment As ISegment
  Dim pStartPoint As IPoint
  Set pStartPoint = New Point
  Dim pEndPoint As IPoint
  Set pEndPoint = New Point
  Dim pClone As IClone
  Dim booFoundCurve As Boolean
  
  Dim lngPathCount As Long
  Dim lngNumVertices As Long
  Dim pRatio As Double
  Dim anIndex As Long
  Dim anIndex2 As Long
  Dim anIndex3 As Long
  Dim lngSegCount As Long
  
  Dim SegCollection As ISegmentCollection
  Dim pNewSegCollection As ISegmentCollection
  
  Dim pPathSegColl As ISegmentCollection
  Dim pNewSegment As ISegment
  Dim pNewLine As esriGeometry.ILine
  
  Dim pNewPolyGeoColl As IGeometryCollection
  Set pNewPolyGeoColl = New Polyline
  Dim pNewPathGeometry As IGeometry
  
  Dim pRing As IRing
  
  lngPathCount = pGeometryCollection.GeometryCount - 1
  For anIndex = 0 To lngPathCount
    If TypeOf pOrigGeometry Is IPolygon Then
      Set pRing = pGeometryCollection.Geometry(anIndex)
      Set pPath = pRing
    Else
      Set pPath = pGeometryCollection.Geometry(anIndex)
    End If
    Set SegCollection = pPath
    Set pNewSegCollection = New Path
    lngSegCount = SegCollection.SegmentCount - 1
    For anIndex2 = 0 To lngSegCount
      Set pSegment = SegCollection.Segment(anIndex2)
      Set pGeometry = pSegment
      If pGeometry.GeometryType <> esriGeometryLine Then ' IF SEGMENT IS CURVE
        booFoundCurve = True
        lngNumVertices = Int((pSegment.length / dblFullLength) * NumVertices)
        If lngNumVertices < 8 Then lngNumVertices = 8
        pRatio = 1 / lngNumVertices
       
        Set pCurve = pSegment
        Set pPathSegColl = New Path
        Set pNewSegment = New esriGeometry.Line
        Set pStartPoint = pCurve.FromPoint
        For anIndex3 = 1 To lngNumVertices
          pCurve.QueryPoint 0, (pRatio * anIndex3), True, pEndPoint
          pNewSegment.FromPoint = pStartPoint
          pNewSegment.ToPoint = pEndPoint
          
          Set pClone = pNewSegment
          pPathSegColl.AddSegment pClone.Clone
          Set pClone = pEndPoint
          Set pStartPoint = pClone.Clone
        Next anIndex3
        pNewSegCollection.AddSegmentCollection pPathSegColl

      Else      ' IF SEGMENT IS ACTUALLY LINE, DON'T ADD MIDPOINTS
        Set pClone = pSegment
        pNewSegCollection.AddSegment pClone.Clone
      End If
    Next anIndex2
    Set pNewPathGeometry = pNewSegCollection
    pNewPolyGeoColl.AddGeometry pNewPathGeometry

  Next anIndex

  
  Dim pNewPolyline As IPolyline
  
  If booFoundCurve Or (TypeOf pOrigGeometry Is IPolygon) Then
  
    Set pNewPolyline = pNewPolyGeoColl
    Dim pTopoOp As ITopologicalOperator
    Set pTopoOp = pNewPolyline
    pTopoOp.Simplify
    Set pNewPolyline.SpatialReference = pSpRef

  Else
    Set pNewPolyline = pOrigGeometry
    Set pNewPolyline.SpatialReference = pSpRef
  End If
    
  Set CurveToPolyline = pNewPolyline
  Exit Function
  
erh:
    MsgBox "Failed in CurveToPolyline: " & vbCrLf & "Error = " & err.Description & vbCrLf & "Line Number = " & CStr(Erl)
End Function
Public Sub Graphic_MakeFromGeometry(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, Optional strName As String)
  
  Dim pMxDocument As esriArcMapUI.IMxDocument
  Dim pActiveView As esriCarto.IActiveView
  
  Dim pGContainer As IGraphicsContainer
  Set pGContainer = pMxDoc.FocusMap
  
  Dim pElement As IElement
  Dim pPolygonElement As IPolygonElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties
  
  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone
  
  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType
  
  'ADD GEOMETRY, NAME AND SPATIAL REFERENCE TO GRAPHIC ELEMENT
  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
    Case 4, 11:
      Set pElement = New PolygonElement
    Case 5:
      Set pElement = New RectangleElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
  End Select
    
  pElement.Geometry = pNewGeometry
  Set pGraphicElement = pElement
  Set pSpatialReference = pGeometry.SpatialReference
  Set pGraphicElement.SpatialReference = pSpatialReference
  Set pElementProperties = pElement
  pElementProperties.Name = strName

  ' ADD GRAPHIC TO GRAPHICS CONTAINER
  pGContainer.AddElement pElement, 0

  'Draw
  pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing

End Sub
Public Function Graphic_ReturnElementFromGeometry(ByRef pMxDoc As IMxDocument, ByRef pGeometry As IGeometry, _
    Optional strName As String, Optional AddToView As Boolean) As IElement
  
  Dim pMxDocument As esriArcMapUI.IMxDocument
  Dim pActiveView As esriCarto.IActiveView
  
  Dim pGContainer As IGraphicsContainer
  Set pGContainer = pMxDoc.FocusMap
  
  Dim pElement As IElement
  Dim pPolygonElement As IPolygonElement
  Dim pSpatialReference As ISpatialReference
  Dim pGraphicElement As IGraphicElement
  Dim pElementProperties As IElementProperties
  
  Dim pClone As IClone
  Set pClone = pGeometry
  Dim pNewGeometry As IGeometry
  Set pNewGeometry = pClone.Clone
  
  Dim pGeometryType As esriGeometryType
  pGeometryType = pNewGeometry.GeometryType
  
  'ADD GEOMETRY, NAME AND SPATIAL REFERENCE TO GRAPHIC ELEMENT
  Select Case pGeometryType
    Case 0:
      MsgBox "Null Geometry!  No graphic added..."
    Case 1:
      Set pElement = New MarkerElement
    Case 3, 6, 13, 14, 15, 16:
      Set pElement = New LineElement
    Case 4, 11:
      Set pElement = New PolygonElement
    Case 5:
      Set pElement = New RectangleElement
    Case Else:
      MsgBox "Unexpected Shape Type:  Unable to add graphic..."
      Exit Function
  End Select
  
  pElement.Geometry = pNewGeometry
  Set pGraphicElement = pElement
  Set pSpatialReference = pGeometry.SpatialReference
  Set pGraphicElement.SpatialReference = pSpatialReference
  Set pElementProperties = pElement
  pElementProperties.Name = strName
  
  If AddToView Then
    ' ADD GRAPHIC TO GRAPHICS CONTAINER
    pGContainer.AddElement pElement, 0
    'Draw
    pMxDoc.ActiveView.PartialRefresh esriViewGraphics, Nothing, Nothing
  End If
  
  Set Graphic_ReturnElementFromGeometry = pElement

End Function

Public Sub ShowVertices(pMxDoc As IMxDocument, pGeometry As IGeometry, Optional strName As String, _
        Optional DeleteCurrentGraphicsWithName As Boolean)

  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  Dim pPoly As IPolygon
  Dim pLine As IPolyline
  Dim pOutVertex As IPoint, lOutPart As Long, lOutVertex As Long
  
  If DeleteCurrentGraphicsWithName And (strName <> "") Then
    Call DeleteGraphicsByName(pMxDoc, "DeleteMe")
  End If
  
  Dim pArray As IArray
  Set pArray = New esriSystem.Array
  
  Dim pPointCollection As IPointCollection
  Set pPointCollection = pGeometry
  
  Dim pPointEnum As IEnumVertex
  Set pPointEnum = pPointCollection.EnumVertices
  
  pPointEnum.Reset
  
  Dim pVertex As IPoint
  Set pVertex = New Point
  'Query the next vertex - have to cocreate the point
  'QueryNext is faster than Next, because the method doesn't have
  'to create the point each time
  pPointEnum.QueryNext pVertex, lOutPart, lOutVertex
  
  Do While Not pVertex.IsEmpty
    Graphic_MakeFromGeometry pMxDoc, pVertex, strName
    pPointEnum.QueryNext pVertex, lOutPart, lOutVertex
'    Debug.Print lOutPart & ",    " & lOutVertex & ",  " & pVertex.IsEmpty
  Loop

End Sub
Public Function CalcBearing(ByRef Point1 As IPoint, ByRef Point2 As IPoint) As Double

  Dim dblBearing As Double

  Dim xDist As Double
  Dim yDist As Double
  Dim xyTanDeg As Double
  
  xDist = (Point1.x - Point2.x)
  yDist = (Point1.Y - Point2.Y)
  If yDist = 0 Then
    If xDist < 0 Then
      xyTanDeg = -90
    ElseIf xDist = 0 Then
      xyTanDeg = 0
    Else
      xyTanDeg = 90
    End If
  Else
    xyTanDeg = AsDegrees(Atn(xDist / yDist))
  End If

  If (yDist >= 0) Then
    dblBearing = 180 + xyTanDeg
  Else
    If (xDist <= 0) Then
      dblBearing = xyTanDeg
    Else
      dblBearing = 360 + xyTanDeg
    End If
  End If ' END CALCULATING BEARING
  
  dblBearing = Abs(dblBearing)
  CalcBearing = dblBearing

End Function

Public Function CalcBearing2(ByRef Point1 As IPoint, ByRef Point2 As IPoint) As Double

  Dim dblBearing As Double

  Dim xDist As Double
  Dim yDist As Double
  Dim xyTanDeg As Double
  
  xDist = (Point1.x - Point2.x)
  yDist = (Point1.Y - Point2.Y)
  
  If xDist = 0 And yDist = 0 Then
    CalcBearing2 = -9999
  Else
    If yDist = 0 Then
      If xDist < 0 Then
        xyTanDeg = -90
      ElseIf xDist = 0 Then
        xyTanDeg = 0
      Else
        xyTanDeg = 90
      End If
    Else
      xyTanDeg = AsDegrees(Atn(xDist / yDist))
    End If
  
    If (yDist >= 0) Then
      dblBearing = 180 + xyTanDeg
    Else
      If (xDist <= 0) Then
        dblBearing = xyTanDeg
      Else
        dblBearing = 360 + xyTanDeg
      End If
    End If ' END CALCULATING BEARING
    
    dblBearing = Abs(dblBearing)
    CalcBearing2 = dblBearing
  End If

End Function

Public Function CalcDistMatrix(pArray As esriSystem.IArray, Optional IncludeLine As Boolean, _
    Optional IncludeBearing As Boolean, Optional pApp As IApplication) As Collection

  Screen.MousePointer = vbHourglass
  
  ' RETURNS A COLLECTION OF IVariantArray OBJECTS
  ' EACH IVariantArray IDENTIFIED BY STRING; CONCATENATION OF [ORIGIN INDEX] & "_" & [DESTINATION INDEX]
  ' EACH IVariantArray OBJECT CONTAINS:
  '       0) ORIGIN SHAPE INDEX VALUE
  '       1) DESTINATION SHAPE INDEX VALUE
  '       2) DISTANCE
  '       3) CONNECTOR POLYLINE:  OPTIONAL; CONTAINS BOOLEAN "FALSE" IF NOT REQUESTED
  '       4) BEARING:             OPTIONAL; CONTAINS BOOLEAN "FALSE" IF NOT REQUESTED
  
  Dim pCollection As Collection
  Set pCollection = New Collection
  
  Dim pProxOp As IProximityOperator
  
  Dim lngIndex As Long
  Dim lngIndex2 As Long
  
  Dim lngArrayMaxIndex As Long
  lngArrayMaxIndex = pArray.Count - 1
  
  Dim pGeometry1 As IGeometry
  Dim pGeometry2 As IGeometry
  
  Dim dblDistance As Double
  Dim dblBearing As Double
  Dim dblRevBearing As Double
  Dim strID As String
  Dim strRevID As String
  Dim pConnector As IPolyline
  Dim pRevConnector As IPolyline
  
  Dim pVarArray As IVariantArray
  Dim pRevVarArray As IVariantArray
  
  If Not pApp Is Nothing Then
      ' PROGRESS BAR STUFF
    Dim pSBar As IStatusBar
    Set pSBar = pApp.StatusBar
    Dim pPro As IStepProgressor
    Set pPro = pSBar.ProgressBar
    Dim lngCounter As Long
    lngCounter = 0
    Dim lngTotalCount As Long
    lngTotalCount = (((lngArrayMaxIndex + 1) * lngArrayMaxIndex) / 2)
    Dim strTotalCount As String
    strTotalCount = CStr(lngTotalCount)
    pPro.position = 1
    pSBar.ShowProgressBar "Building Preliminary Distance Matrix:  Step 1 of " & strTotalCount & "...", 1, _
            lngTotalCount, 1, True
  End If
  
  Dim pOutputPointCollection As IPointCollection
  
  For lngIndex = 0 To lngArrayMaxIndex
    Set pGeometry1 = pArray.Element(lngIndex)
    
    For lngIndex2 = lngIndex To lngArrayMaxIndex
      
      strID = CStr(lngIndex) & "_" & CStr(lngIndex2)
      strRevID = CStr(lngIndex2) & "_" & CStr(lngIndex)
      
      Set pVarArray = New varArray
      Set pRevVarArray = New varArray
      
      ' FIRST ELEMENT
      pVarArray.Add lngIndex         ' FIRST VALUES IN THE ARRAY ARE ORIGIN NODES
      pRevVarArray.Add lngIndex2
      
      ' SECOND ELEMENT
      pVarArray.Add lngIndex2        ' SECOND VALUES IN THE ARRAY ARE "TO" NODES
      pRevVarArray.Add lngIndex
      
      If lngIndex = lngIndex2 Then   ' IF MEASURING DISTANCE TO ITSELF
        Set pConnector = New Polyline
        pConnector.SetEmpty
        Set pRevConnector = New Polyline
        pRevConnector.SetEmpty
        
        ' THIRD ELEMENT
        pVarArray.Add 0              ' THIRD VALUE IS DISTANCE
        
        ' FOURTH ELEMENT
        If IncludeLine Then
          pVarArray.Add pConnector   ' FOURTH VALUE IS CONNECTION LINE
        Else
          pVarArray.Add False        ' FOURTH VALUE:  JUST ADDING SMALL PLACEHOLDER ELEMENT
        End If
        
        ' FIFTH ELEMENT
        If IncludeBearing Then
          pVarArray.Add -999         ' FIFTH VALUE IS BEARING
        Else
          pVarArray.Add False
        End If
        
        ' ADD VARARRAY TO ORIGINAL COLLECTION
        pCollection.Add pVarArray, strID
      
      Else
        
        If Not pApp Is Nothing Then
          lngCounter = lngCounter + 1
          pPro.Message = "Building Preliminary Distance Matrix:  Step " & CStr(lngCounter) & " of " & strTotalCount & "..."
          pSBar.StepProgressBar
        End If
        
        Set pGeometry2 = pArray.Element(lngIndex2)
        
        If IncludeLine Or IncludeBearing Then
          Dim pLineArray As IArray
          Set pLineArray = CalcClosestPoints(pGeometry1, pGeometry2, 10)
          
          If TypeOf pLineArray.Element(0) Is esriSystem.IStringArray Then      ' FUNCTION FAILED FOR SOME REASON
            Dim pStrArray As IStringArray
            Set pStrArray = pLineArray.Element(0)
            MsgBox "Failed to connect:" & vbCrLf & "Message = " & pStrArray.Element(0) & vbCrLf & _
                   "Index 1 = " & CStr(lngIndex) & " of " & CStr(lngArrayMaxIndex) & vbCrLf & _
                   "Index 2 = " & CStr(lngIndex2) & " of " & CStr(lngArrayMaxIndex)
            Set pConnector = New Polyline
            pConnector.SetEmpty
            Set pRevConnector = New Polyline
            pRevConnector.SetEmpty
            
            ' THIRD ELEMENT
            pVarArray.Add 0              ' THIRD VALUE IS DISTANCE
            
            ' FOURTH ELEMENT
            If IncludeLine Then
              pVarArray.Add pConnector   ' FOURTH VALUE IS CONNECTION LINE
            Else
              pVarArray.Add False        ' FOURTH VALUE:  JUST ADDING SMALL PLACEHOLDER ELEMENT
            End If
            
            ' FIFTH ELEMENT
            If IncludeBearing Then
              pVarArray.Add -999         ' FIFTH VALUE IS BEARING
            Else
              pVarArray.Add False
            End If
          Else
          
            If IncludeLine Then
              Set pConnector = pLineArray.Element(0)
              Set pRevConnector = New Polyline
              Set pOutputPointCollection = pRevConnector
              pOutputPointCollection.AddPoint pLineArray.Element(2)
              pOutputPointCollection.AddPoint pLineArray.Element(1)
              
              ' THIRD ELEMENT
              pVarArray.Add pConnector.length
              pRevVarArray.Add pConnector.length
              
              ' FOURTH ELEMENT
              pVarArray.Add pConnector
              pRevVarArray.Add pRevConnector
            Else
              ' FOURTH ELEMENT
              pVarArray.Add False        ' JUST ADDING SMALL PLACEHOLDER ELEMENT
              pRevVarArray.Add False     ' JUST ADDING SMALL PLACEHOLDER ELEMENT
            End If
            If IncludeBearing Then
              dblBearing = CalcBearing(pLineArray.Element(1), pLineArray.Element(2))
              If dblBearing < 180 Then
                dblRevBearing = dblBearing + 180
              Else
                dblRevBearing = dblBearing - 180
              End If
              ' FIFTH ELEMENT
              pVarArray.Add dblBearing
              pRevVarArray.Add dblRevBearing
            Else
              ' FIFTH ELEMENT
              pVarArray.Add False
              pRevVarArray.Add False
            End If
          End If
        Else
          ' THIRD ELEMENT
          Set pProxOp = pGeometry1
          dblDistance = pProxOp.ReturnDistance(pGeometry2)
          pVarArray.Add dblDistance
          pRevVarArray.Add dblDistance
          ' FOURTH ELEMENT   (DISTANCE)
          pVarArray.Add False
          pRevVarArray.Add False
          ' FIFTH ELEMENT   (BEARING)
          pVarArray.Add False
          pRevVarArray.Add False
        End If
        
        pCollection.Add pVarArray, strID
        pCollection.Add pRevVarArray, strRevID

      End If
    Next lngIndex2
  Next lngIndex
  
  Set CalcDistMatrix = pCollection

  Screen.MousePointer = vbDefault
  If Not pApp Is Nothing Then
    pPro.position = 1
    pSBar.HideProgressBar
  End If

End Function
Public Function CalcClosestPoints(ByVal Shape1 As IGeometry, ByVal shape2 As IGeometry, Optional intMaxCurveRepeat As Integer) As IArray

' CalcClosestPoints
' Jenness Enterprises (www.jennessent.com)
' Given two shapes, this script returns an IARRAY object containing the line connecting the closest points on each shape, plus the connection points
' CURRENTLY DOES NOT GUARANTEE SUCCESS WITH TRUE CURVES BECAUSE VERTICES ARE NOT GOOD QUERY POINTS; ATTEMPTS SEVERAL RUNS BACK AND FORTH

' Dim pRelationalOperator As IRelationalOperator
Dim pGeometryType1 As esriGeometryType
Dim pGeometryType2 As esriGeometryType
Dim pGeometry1 As IGeometry
Dim pGeometry2 As IGeometry

Set pGeometry1 = Shape1
Set pGeometry2 = shape2

pGeometryType2 = shape2.GeometryType

' IF SHAPE #2 HAPPENS TO BE POINT, SET THAT ONE FIRST
Dim ShouldReverse As Boolean
ShouldReverse = False
If pGeometryType2 = esriGeometryPoint Then
  Set pGeometry1 = shape2
  Set pGeometry2 = Shape1
  pGeometryType2 = pGeometry2.GeometryType
  ShouldReverse = True
End If
  
pGeometryType1 = pGeometry1.GeometryType

Dim pArray As IArray
Set pArray = New esriSystem.Array
Dim pOutputLine As IPolyline
Set pOutputLine = New Polyline
Dim pOutputPointCollection As IPointCollection
Set pOutputPointCollection = pOutputLine

Dim pStartPoint As IPoint
Dim pEndPoint As IPoint

Dim pPoint1 As IPoint
Dim pPoint2 As IPoint
Set pPoint2 = New Point

Dim pProximityOp As IProximityOperator
Dim pStringArray As IStringArray

' CHECK FOR NULL SHAPES
' CHECK FOR INTERSECTING SHAPES
' NOT SURE IF THIS WILL WORK WITH MULTIPOINTS

If pGeometry1.IsEmpty Or pGeometry2.IsEmpty Then
  Set pStringArray = New strArray
  pStringArray.Add "Empty Shapes"
  pStringArray.Add CStr(pGeometry1.IsEmpty)
  pStringArray.Add CStr(pGeometry2.IsEmpty)
  pArray.Add pStringArray
  Set CalcClosestPoints = pArray
  Exit Function
Else
  Set pProximityOp = pGeometry1
  If pProximityOp.ReturnDistance(pGeometry2) = 0 Then
    Set pStringArray = New strArray
    pStringArray.Add "Intersecting Shapes"
    pArray.Add pStringArray
    Set CalcClosestPoints = pArray
    Exit Function
  End If
End If

If pGeometryType1 = esriGeometryPoint Then
  Set pPoint1 = pGeometry1
  
  If pGeometryType2 = esriGeometryPoint Then
    Set pPoint2 = pGeometry2
    
    If pPoint1.x = pPoint2.x And pPoint1.Y = pPoint2.Y Then
      If ShouldReverse Then
        pArray.Add pOutputLine
        pArray.Add pPoint2
        pArray.Add pPoint1
      Else
        pArray.Add pOutputLine
        pArray.Add pPoint1
        pArray.Add pPoint2
      End If
    Else
      If ShouldReverse Then
        pOutputPointCollection.AddPoint pPoint2
        pOutputPointCollection.AddPoint pPoint1
        pArray.Add pOutputLine
        pArray.Add pPoint2
        pArray.Add pPoint1
      Else
        pOutputPointCollection.AddPoint pPoint1
        pOutputPointCollection.AddPoint pPoint2
        pArray.Add pOutputLine
        pArray.Add pPoint1
        pArray.Add pPoint2
      End If
    End If
  Else
    
    Set pProximityOp = pGeometry2
    
    pProximityOp.QueryNearestPoint pPoint1, esriNoExtension, pPoint2
    
    If ShouldReverse Then
      pOutputPointCollection.AddPoint pPoint2
      pOutputPointCollection.AddPoint pPoint1
      pArray.Add pOutputLine
      pArray.Add pPoint2
      pArray.Add pPoint1
    Else
      pOutputPointCollection.AddPoint pPoint1
      pOutputPointCollection.AddPoint pPoint2
      pArray.Add pOutputLine
      pArray.Add pPoint1
      pArray.Add pPoint2
    End If
    
    
  End If
Else
  Dim dblTestDistance As Double
  Dim pEnvelope As IEnvelope
  Dim pEnvelope2 As IEnvelope
  Set pEnvelope = pGeometry1.Envelope
  Set pEnvelope2 = pGeometry2.Envelope
  pEnvelope.Union pEnvelope2
  If pEnvelope.Height > pEnvelope.Width Then
    dblTestDistance = pEnvelope.Height * 2
  Else
    dblTestDistance = pEnvelope.Width * 2
  End If
  'dblTestDistance = (pEnvelope.Height * pEnvelope.Width)
  Dim dblMaxDistance As Double
  dblMaxDistance = dblTestDistance
  
  Dim pPointCollection1 As IPointCollection
  Dim pPointCollection2 As IPointCollection
  
  If pGeometry1.GeometryType = esriGeometryEnvelope Then
    Dim pTempEnv As IEnvelope
    Set pTempEnv = pGeometry1
    Dim pTempPoly1 As IPolygon
    Dim pTempPoint1 As IPoint
    Set pTempPoly1 = New Polygon
    Set pPointCollection1 = pTempPoly1
    Dim dXmin1 As Double
    Dim dYmin1 As Double
    Dim dXmax1 As Double
    Dim dYmax1 As Double
    pTempEnv.QueryCoords dXmin1, dYmin1, dXmax1, dYmax1
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmin1
    pTempPoint1.Y = dYmin1
    pPointCollection1.AddPoint pTempPoint1
    
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmin1
    pTempPoint1.Y = dYmax1
    pPointCollection1.AddPoint pTempPoint1
    
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmax1
    pTempPoint1.Y = dYmax1
    pPointCollection1.AddPoint pTempPoint1
    
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmax1
    pTempPoint1.Y = dYmin1
    pPointCollection1.AddPoint pTempPoint1
  Else
    Set pPointCollection1 = pGeometry1
  End If
  
  If pGeometry2.GeometryType = esriGeometryEnvelope Then
    Dim pTempEnv2 As IEnvelope
    Set pTempEnv2 = pGeometry2
    Dim pTempPoly2 As IPolygon
    Dim pTempPoint2 As IPoint
    Set pTempPoly2 = New Polygon
    Set pPointCollection2 = pTempPoly2
    Dim dXmin2 As Double
    Dim dYmin2 As Double
    Dim dXmax2 As Double
    Dim dYmax2 As Double
    pTempEnv2.QueryCoords dXmin2, dYmin2, dXmax2, dYmax2
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmin2
    pTempPoint2.Y = dYmin2
    pPointCollection2.AddPoint pTempPoint2
    
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmin2
    pTempPoint2.Y = dYmax2
    pPointCollection2.AddPoint pTempPoint2
    
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmax2
    pTempPoint2.Y = dYmax2
    pPointCollection2.AddPoint pTempPoint2
    
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmax2
    pTempPoint2.Y = dYmin2
    pPointCollection2.AddPoint pTempPoint2
  Else
    Set pPointCollection2 = pGeometry2
  End If
  
  Dim pClone As IClone
  
  Dim pVertex As IPoint
  Set pVertex = New Point
  
  Dim pPointEnum As IEnumVertex
  Dim lngOutPart As Long
  Dim lngOutVertex As Long
  
  Set pPointEnum = pPointCollection1.EnumVertices
  pPointEnum.Reset
  pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  
  ' CHECK IF CURVES; THIS CODE JUST CHECKS FIRST SEGMENT FOR CURVATURE
  Dim booWorkingWithCurves As Boolean
  Dim pSegmentCollection1 As ISegmentCollection
  Set pSegmentCollection1 = pGeometry1
  Dim pSegment1 As ISegment
  Set pSegment1 = pSegmentCollection1.Segment(0)
  Dim pGeometryTypeA As esriGeometryType
  pGeometryTypeA = pSegment1.GeometryType
  
  Dim pSegmentCollection2 As ISegmentCollection
  Set pSegmentCollection2 = pGeometry2
  Dim pSegment2 As ISegment
  Set pSegment2 = pSegmentCollection2.Segment(0)
  Dim pGeometryTypeB As esriGeometryType
  pGeometryTypeB = pSegment2.GeometryType
  
  booWorkingWithCurves = (pGeometryTypeA = esriGeometryBezier3Curve) Or _
    (pGeometryTypeA = esriGeometryCircularArc) Or _
    (pGeometryTypeA = esriGeometryEllipticArc) Or _
    (pGeometryTypeB = esriGeometryBezier3Curve) Or _
    (pGeometryTypeB = esriGeometryCircularArc) Or _
    (pGeometryTypeB = esriGeometryEllipticArc)

  Do While Not pVertex.IsEmpty
    Set pProximityOp = pGeometry2
    dblTestDistance = pProximityOp.ReturnDistance(pVertex)
    If dblTestDistance < dblMaxDistance Then
      dblMaxDistance = dblTestDistance
      Set pClone = pVertex
      Set pPoint1 = pClone.Clone
      pProximityOp.QueryNearestPoint pVertex, esriNoExtension, pPoint2
    End If
    pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  Loop
  
  Set pPointEnum = pPointCollection2.EnumVertices
  pPointEnum.Reset
  pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  
  Do While Not pVertex.IsEmpty
    Set pProximityOp = pGeometry1
    dblTestDistance = pProximityOp.ReturnDistance(pVertex)
    If dblTestDistance < dblMaxDistance Then
      dblMaxDistance = dblTestDistance
      Set pClone = pVertex
      Set pPoint2 = pClone.Clone
      pProximityOp.QueryNearestPoint pVertex, esriNoExtension, pPoint1
    End If
    pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  Loop
  
  ' FOR DEBUGGING
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
  ' IF WORKING WITH CURVES, GO BACK AND FORTH A FEW TIMES
  If booWorkingWithCurves Then
    Dim intRepeat As Integer
    Dim pPoint1Temp As IPoint, pPoint2Temp As IPoint
    
    Do Until (intRepeat = intMaxCurveRepeat)
 '     Graphic_MakeFromGeometry pMxDoc, pPoint1, "DeleteMe"
 '     Graphic_MakeFromGeometry pMxDoc, pPoint2, "DeleteMe"
      
      Set pProximityOp = pGeometry2
      pProximityOp.QueryNearestPoint pPoint1, esriNoExtension, pPoint2
      
      Set pProximityOp = pGeometry1
      pProximityOp.QueryNearestPoint pPoint2, esriNoExtension, pPoint1
      
      intRepeat = intRepeat + 1
    
    Loop
  
  End If
  
  If ShouldReverse Then
    pOutputPointCollection.AddPoint pPoint2
    pOutputPointCollection.AddPoint pPoint1
    pArray.Add pOutputLine
    pArray.Add pPoint2
    pArray.Add pPoint1
  Else
    pOutputPointCollection.AddPoint pPoint1
    pOutputPointCollection.AddPoint pPoint2
    pArray.Add pOutputLine
    pArray.Add pPoint1
    pArray.Add pPoint2
  End If
  
End If

Set CalcClosestPoints = pArray

End Function
Public Function CalcClosestPoints2(ByVal Shape1 As IGeometry, ByVal shape2 As IGeometry, Optional intMaxCurveRepeat As Integer) As IArray

' CalcClosestPoints2
' Jenness Enterprises (www.jennessent.com)
' Given two shapes, this script returns an IARRAY object containing the line connecting the closest points on each shape, plus the connection points
' CURRENTLY DOES NOT GUARANTEE SUCCESS WITH TRUE CURVES BECAUSE VERTICES ARE NOT GOOD QUERY POINTS; ATTEMPTS SEVERAL RUNS BACK AND FORTH

' Dim pRelationalOperator As IRelationalOperator
Dim pGeometryType1 As esriGeometryType
Dim pGeometryType2 As esriGeometryType
Dim pGeometry1 As IGeometry
Dim pGeometry2 As IGeometry

Set pGeometry1 = Shape1
Set pGeometry2 = shape2

pGeometryType2 = shape2.GeometryType

' IF SHAPE #2 HAPPENS TO BE POINT, SET THAT ONE FIRST
Dim ShouldReverse As Boolean
ShouldReverse = False
If pGeometryType2 = esriGeometryPoint Then
  Set pGeometry1 = shape2
  Set pGeometry2 = Shape1
  pGeometryType2 = pGeometry2.GeometryType
  ShouldReverse = True
End If
  
pGeometryType1 = pGeometry1.GeometryType

Dim pArray As IArray
Set pArray = New esriSystem.Array
Dim pOutputLine As IPolyline
Set pOutputLine = New Polyline
Dim pOutputPointCollection As IPointCollection
Set pOutputPointCollection = pOutputLine

Dim pStartPoint As IPoint
Dim pEndPoint As IPoint

Dim pPoint1 As IPoint
Dim pPoint2 As IPoint
Set pPoint2 = New Point

Dim pProximityOp As IProximityOperator
Dim pStringArray As IStringArray

' CHECK FOR NULL SHAPES
' CHECK FOR INTERSECTING SHAPES
' NOT SURE IF THIS WILL WORK WITH MULTIPOINTS

If pGeometry1.IsEmpty Or pGeometry2.IsEmpty Then
  Set pStringArray = New strArray
  pStringArray.Add "Empty Shapes"
  pStringArray.Add CStr(pGeometry1.IsEmpty)
  pStringArray.Add CStr(pGeometry2.IsEmpty)
  pArray.Add pStringArray
  Set CalcClosestPoints2 = pArray
  Exit Function
Else
  Set pProximityOp = pGeometry1
  If pProximityOp.ReturnDistance(pGeometry2) = 0 Then
    Set pStringArray = New strArray
    pStringArray.Add "Intersecting Shapes"
    pArray.Add pStringArray
    Set CalcClosestPoints2 = pArray
    Exit Function
  End If
End If

If pGeometryType1 = esriGeometryPoint Then
  Set pPoint1 = pGeometry1
  
  If pGeometryType2 = esriGeometryPoint Then
    Set pPoint2 = pGeometry2
    
    If pPoint1.x = pPoint2.x And pPoint1.Y = pPoint2.Y Then
      If ShouldReverse Then
        pArray.Add pOutputLine
        pArray.Add pPoint2
        pArray.Add pPoint1
      Else
        pArray.Add pOutputLine
        pArray.Add pPoint1
        pArray.Add pPoint2
      End If
    Else
      If ShouldReverse Then
        pOutputPointCollection.AddPoint pPoint2
        pOutputPointCollection.AddPoint pPoint1
        pArray.Add pOutputLine
        pArray.Add pPoint2
        pArray.Add pPoint1
      Else
        pOutputPointCollection.AddPoint pPoint1
        pOutputPointCollection.AddPoint pPoint2
        pArray.Add pOutputLine
        pArray.Add pPoint1
        pArray.Add pPoint2
      End If
    End If
  Else
    
    Set pProximityOp = pGeometry2
    
    pProximityOp.QueryNearestPoint pPoint1, esriNoExtension, pPoint2
    
    If ShouldReverse Then
      pOutputPointCollection.AddPoint pPoint2
      pOutputPointCollection.AddPoint pPoint1
      pArray.Add pOutputLine
      pArray.Add pPoint2
      pArray.Add pPoint1
    Else
      pOutputPointCollection.AddPoint pPoint1
      pOutputPointCollection.AddPoint pPoint2
      pArray.Add pOutputLine
      pArray.Add pPoint1
      pArray.Add pPoint2
    End If
    
    
  End If
Else
  Dim dblTestDistance As Double
  Dim pEnvelope As IEnvelope
  Dim pEnvelope2 As IEnvelope
  Set pEnvelope = pGeometry1.Envelope
  Set pEnvelope2 = pGeometry2.Envelope
  pEnvelope.Union pEnvelope2
  If pEnvelope.Height > pEnvelope.Width Then
    dblTestDistance = pEnvelope.Height * 2
  Else
    dblTestDistance = pEnvelope.Width * 2
  End If
  'dblTestDistance = (pEnvelope.Height * pEnvelope.Width)
  Dim dblMaxDistance As Double
  dblMaxDistance = dblTestDistance
  
  Dim pPointCollection1 As IPointCollection
  Dim pPointCollection2 As IPointCollection
  
  If pGeometry1.GeometryType = esriGeometryEnvelope Then
    Dim pTempEnv As IEnvelope
    Set pTempEnv = pGeometry1
    Dim pTempPoly1 As IPolygon
    Dim pTempPoint1 As IPoint
    Set pTempPoly1 = New Polygon
    Set pPointCollection1 = pTempPoly1
    Dim dXmin1 As Double
    Dim dYmin1 As Double
    Dim dXmax1 As Double
    Dim dYmax1 As Double
    pTempEnv.QueryCoords dXmin1, dYmin1, dXmax1, dYmax1
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmin1
    pTempPoint1.Y = dYmin1
    pPointCollection1.AddPoint pTempPoint1
    
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmin1
    pTempPoint1.Y = dYmax1
    pPointCollection1.AddPoint pTempPoint1
    
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmax1
    pTempPoint1.Y = dYmax1
    pPointCollection1.AddPoint pTempPoint1
    
    Set pTempPoint1 = New Point
    pTempPoint1.x = dXmax1
    pTempPoint1.Y = dYmin1
    pPointCollection1.AddPoint pTempPoint1
  Else
    Set pPointCollection1 = pGeometry1
  End If
  
  If pGeometry2.GeometryType = esriGeometryEnvelope Then
    Dim pTempEnv2 As IEnvelope
    Set pTempEnv2 = pGeometry2
    Dim pTempPoly2 As IPolygon
    Dim pTempPoint2 As IPoint
    Set pTempPoly2 = New Polygon
    Set pPointCollection2 = pTempPoly2
    Dim dXmin2 As Double
    Dim dYmin2 As Double
    Dim dXmax2 As Double
    Dim dYmax2 As Double
    pTempEnv2.QueryCoords dXmin2, dYmin2, dXmax2, dYmax2
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmin2
    pTempPoint2.Y = dYmin2
    pPointCollection2.AddPoint pTempPoint2
    
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmin2
    pTempPoint2.Y = dYmax2
    pPointCollection2.AddPoint pTempPoint2
    
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmax2
    pTempPoint2.Y = dYmax2
    pPointCollection2.AddPoint pTempPoint2
    
    Set pTempPoint2 = New Point
    pTempPoint2.x = dXmax2
    pTempPoint2.Y = dYmin2
    pPointCollection2.AddPoint pTempPoint2
  Else
    Set pPointCollection2 = pGeometry2
  End If
  
  Dim pClone As IClone
  
  Dim pVertex As IPoint
  Set pVertex = New Point
  
  Dim pPointEnum As IEnumVertex
  Dim lngOutPart As Long
  Dim lngOutVertex As Long
  
  Set pPointEnum = pPointCollection1.EnumVertices
  pPointEnum.Reset
  pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  
  ' CHECK IF CURVES; THIS CODE JUST CHECKS FIRST SEGMENT FOR CURVATURE
  Dim booWorkingWithCurves As Boolean
  Dim pSegmentCollection1 As ISegmentCollection
  Set pSegmentCollection1 = pGeometry1
  Dim pSegment1 As ISegment
  Set pSegment1 = pSegmentCollection1.Segment(0)
  Dim pGeometryTypeA As esriGeometryType
  pGeometryTypeA = pSegment1.GeometryType
  
  Dim pSegmentCollection2 As ISegmentCollection
  Set pSegmentCollection2 = pGeometry2
  Dim pSegment2 As ISegment
  Set pSegment2 = pSegmentCollection2.Segment(0)
  Dim pGeometryTypeB As esriGeometryType
  pGeometryTypeB = pSegment2.GeometryType
  
  booWorkingWithCurves = (pGeometryTypeA = esriGeometryBezier3Curve) Or _
    (pGeometryTypeA = esriGeometryCircularArc) Or _
    (pGeometryTypeA = esriGeometryEllipticArc) Or _
    (pGeometryTypeB = esriGeometryBezier3Curve) Or _
    (pGeometryTypeB = esriGeometryCircularArc) Or _
    (pGeometryTypeB = esriGeometryEllipticArc)

  Do While Not pVertex.IsEmpty
    Set pProximityOp = pGeometry2
    dblTestDistance = pProximityOp.ReturnDistance(pVertex)
    If dblTestDistance < dblMaxDistance Then
      dblMaxDistance = dblTestDistance
      Set pClone = pVertex
      Set pPoint1 = pClone.Clone
      pProximityOp.QueryNearestPoint pVertex, esriNoExtension, pPoint2
    End If
    pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  Loop
  
  Set pPointEnum = pPointCollection2.EnumVertices
  pPointEnum.Reset
  pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  
  Do While Not pVertex.IsEmpty
    Set pProximityOp = pGeometry1
    dblTestDistance = pProximityOp.ReturnDistance(pVertex)
    If dblTestDistance < dblMaxDistance Then
      dblMaxDistance = dblTestDistance
      Set pClone = pVertex
      Set pPoint2 = pClone.Clone
      pProximityOp.QueryNearestPoint pVertex, esriNoExtension, pPoint1
    End If
    pPointEnum.QueryNext pVertex, lngOutPart, lngOutVertex
  Loop
  
  ' FOR DEBUGGING
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
  ' IF WORKING WITH CURVES, GO BACK AND FORTH A FEW TIMES
  If booWorkingWithCurves Then
    Dim intRepeat As Integer
    Dim pPoint1Temp As IPoint, pPoint2Temp As IPoint
    
    Do Until (intRepeat = intMaxCurveRepeat)
 '     Graphic_MakeFromGeometry pMxDoc, pPoint1, "DeleteMe"
 '     Graphic_MakeFromGeometry pMxDoc, pPoint2, "DeleteMe"
      
      Set pProximityOp = pGeometry2
      pProximityOp.QueryNearestPoint pPoint1, esriNoExtension, pPoint2
      
      Set pProximityOp = pGeometry1
      pProximityOp.QueryNearestPoint pPoint2, esriNoExtension, pPoint1
      
      intRepeat = intRepeat + 1
    
    Loop
  
  End If
  
  If ShouldReverse Then
    pOutputPointCollection.AddPoint pPoint2
    pOutputPointCollection.AddPoint pPoint1
    pArray.Add pOutputLine
    pArray.Add pPoint2
    pArray.Add pPoint1
  Else
    pOutputPointCollection.AddPoint pPoint1
    pOutputPointCollection.AddPoint pPoint2
    pArray.Add pOutputLine
    pArray.Add pPoint1
    pArray.Add pPoint2
  End If
  
End If

Set CalcClosestPoints2 = pArray

End Function


Public Function CalcCheckClockwise(theP As IPoint, theQ As IPoint, theR As IPoint) As Boolean
 
On Error GoTo err
' CalcCheckClockwise
' Jenness Enterprises <www.jennessent.com)>
' Given 3 consecutive points, this scripts calculates whether the third point lies to the right
' (clockwise) or to the left (counter-clockwise) of the line connecting the first point to
' the second point.

' CLOCKWISE IF TRUE
CalcCheckClockwise = ((theQ.x * (theR.Y - theP.Y)) + (theQ.Y * (theP.x - theR.x)) - ((theP.x) * (theR.Y)) _
      + ((theP.Y) * (theR.x)) < 0)
    Exit Function
err:
  MsgBox "Messed up CalcCheckClockwise..."
End Function

Public Function PointAdd(pPointA As IPoint, pPointB As IPoint) As IPoint

  Set PointAdd = New Point
  PointAdd.PutCoords pPointA.x + pPointB.x, pPointA.Y + pPointB.Y

End Function

Public Function PointSubtract(pPointA As IPoint, pPointB As IPoint) As IPoint

  Set PointSubtract = New Point
  PointSubtract.PutCoords pPointA.x - pPointB.x, pPointA.Y - pPointB.Y

End Function

Public Function AsRadians(theDegrees As Double) As Double

  AsRadians = dblPI * (theDegrees / 180)

End Function

Public Function AsDegrees(theRadians As Double) As Double

  AsDegrees = (theRadians * 180) / dblPI

End Function

Public Sub CalcPointLine(ptOrigin As IPoint, theLength As Double, dblAzimuth As Double, ptEndPoint As IPoint, _
    Optional pLine As IPolyline)

' Jenness Enterprises <www.jennessent.com>
' Given an origin point, distance and bearing, this script will return a new point at that distance and bearing, and a line
' connecting that new point to the origin point

'' MAKE SURE AZIMUTH IS BETWEEN 0 AND 360
Dim theAzimuth As Double
theAzimuth = dblAzimuth

Set ptEndPoint = New Point

Do While theAzimuth < 0
  theAzimuth = theAzimuth + 360
Loop
Do While theAzimuth > 360
  theAzimuth = theAzimuth - 360
Loop
'theAzimuth = theAzimuth Mod 360
'
'' NEW SEGMENT AND POINT DISTANCE NORTH/SOUTH AND EAST/WEST BASED ON DISTANCE AND BEARING FROM ORIGIN.
'' THERE ARE EIGHT DIFFERENT POSSIBILITIES:  THE BEARING COULD BE ONE OF THE FOUR CARDINAL DIRECTIONS OR IT
'' COULD BE IN ONE OF THE FOUR QUADRANTS.  THE BEARING IS TREATED DIFFERENTLY IN EACH OF THESE CIRCUMSTANCES.
'' THE CALCULATION TO DETERMINE THE NEW POINT LOCATION IS ESSENTIALLY A MATTER OF TAKING THE SINE OR THE
'' COSINE OF THE ANGLE (AFTER CONVERTING IT TO RADIANS), AND MULTIPLYING THAT SINE OR COSINE BY THE MEASURED
'' DISTANCE.  PLEASE CONTACT THE AUTHOR IF THIS DOESN'T MAKE SENSE, OR IF YOU WOULD LIKE FURTHER EXPLANATION.
Dim NorthSouthDistance As Double
Dim EastWestDistance As Double
Dim EastWest As Integer
Dim NorthSouth As Integer

If theAzimuth = 0 Or theAzimuth = 360 Then
  NorthSouthDistance = theLength
  NorthSouth = 1
  EastWestDistance = 0
  EastWest = 1
ElseIf (theAzimuth = 180) Then
  NorthSouthDistance = theLength
  NorthSouth = -1
  EastWestDistance = 0
  EastWest = 1
ElseIf (theAzimuth = 90) Then
  NorthSouthDistance = 0
  NorthSouth = 1
  EastWestDistance = theLength
  EastWest = 1
ElseIf (theAzimuth = 270) Then
  NorthSouthDistance = 0
  NorthSouth = 1
  EastWestDistance = theLength
  EastWest = -1
ElseIf ((theAzimuth > 0) And (theAzimuth < 90)) Then
  NorthSouthDistance = Cos(AsRadians(theAzimuth)) * theLength
  NorthSouth = 1
  EastWestDistance = Sin(AsRadians(theAzimuth)) * theLength
  EastWest = 1
ElseIf ((theAzimuth > 90) And (theAzimuth < 180)) Then
  NorthSouthDistance = (Sin(AsRadians(theAzimuth - 90))) * theLength
  NorthSouth = -1
  EastWestDistance = (Cos(AsRadians(theAzimuth - 90))) * theLength
  EastWest = 1
ElseIf ((theAzimuth > 180) And (theAzimuth < 270)) Then
  NorthSouthDistance = (Cos(AsRadians(theAzimuth - 180))) * theLength
  NorthSouth = -1
  EastWestDistance = (Sin(AsRadians(theAzimuth - 180))) * theLength
  EastWest = -1
ElseIf ((theAzimuth > 270) And (theAzimuth < 360)) Then
  NorthSouthDistance = (Sin(AsRadians(theAzimuth - 270))) * theLength
  NorthSouth = 1
  EastWestDistance = (Cos(AsRadians(theAzimuth - 270))) * theLength
  EastWest = -1
End If

Dim theMovementNorth As Double
Dim theMovementWest As Double

theMovementNorth = NorthSouthDistance * NorthSouth
theMovementWest = EastWestDistance * EastWest

Dim startX As Double
Dim startY As Double

ptOrigin.QueryCoords startX, startY
ptEndPoint.PutCoords startX + theMovementWest, startY + theMovementNorth

Set ptEndPoint.SpatialReference = ptOrigin.SpatialReference

If Not pLine Is Nothing Then
  Dim pPointColl As IPointCollection
  pLine.SetEmpty
  Set pPointColl = pLine
  pPointColl.AddPoint ptOrigin
  pPointColl.AddPoint ptEndPoint
  Set pLine.SpatialReference = ptOrigin.SpatialReference
End If

End Sub
Public Function EllipticArcToPolygon(SegCollection As ISegmentCollection, NumVertices As Long) As IPolygon4

'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
  
'  Dim pEllArc As IEllipticArc
  Dim pCurve As ICurve
  Dim pGeometry As IGeometry
  
  Dim anIndex As Long
  Dim lngSegCount As Long
  lngSegCount = SegCollection.SegmentCount - 1
  Dim theLength As Double
  theLength = 0
  Dim theTestLength As Double
  Dim lngLengths() As Long
  ReDim lngLengths(lngSegCount)
  For anIndex = 0 To lngSegCount
    theTestLength = SegCollection.Segment(anIndex).length
    theLength = theLength + theTestLength
    lngLengths(anIndex) = theTestLength
  Next anIndex
  
  Dim pProportion As Double
  Dim lngVertices() As Long
  Dim lngNumVertices As Long
  ReDim lngVertices(lngSegCount)
  For anIndex = 0 To lngSegCount
    lngNumVertices = Int((lngLengths(anIndex) / theLength) * NumVertices)
    If lngNumVertices < 8 Then lngNumVertices = 8
    lngVertices(anIndex) = lngNumVertices
  Next anIndex
  
  Dim pMpt As IPointCollection
  Set pMpt = New Multipoint
  Dim pPoint As IPoint
  Set pPoint = New Point
  Dim pClone As IClone
  
  Dim pRatio As Double
  Dim anIndex2 As Long
  
  For anIndex = 0 To lngSegCount
    lngNumVertices = lngVertices(anIndex)
    pRatio = 1 / lngNumVertices
    Set pCurve = SegCollection.Segment(anIndex)
    
    For anIndex2 = 0 To lngNumVertices
'      If pGeometry.GeometryType = esriGeometryEllipticArc Then
      pCurve.QueryPoint 0, (pRatio * anIndex2), True, pPoint
      Set pClone = pPoint
        
 '   Graphic_MakeFromGeometry pMxDoc, pPoint, "DeleteMe"
    
      pMpt.AddPoint pClone.Clone
    Next anIndex2
  Next anIndex
  
  Dim pPoly4 As IPolygon4
  Dim pTopoOp2 As ITopologicalOperator2
  Dim pTopoOp3 As ITopologicalOperator3
  Set pTopoOp2 = pMpt
  Set pPoly4 = pTopoOp2.ConvexHull
  Set pTopoOp3 = pPoly4
  pTopoOp3.IsKnownSimple = False
  pTopoOp3.Simplify
  
  Set EllipticArcToPolygon = pPoly4

End Function


Public Function FeaturePlanetOGraphicToPlanetOCentric(pGeometry As IGeometry, Optional dblMajorAxis As Double = -999, _
    Optional dblMinorAxis As Double = -999, Optional dblLongShift As Double = 0) As IGeometry
  


  ' ASSUMES pGeometry IS IN A GEOGRAPHIC PROJECTION AND IS EITHER A POINT, POLYLINE, POLYGON OR MULTIPOINT
  
  If dblMajorAxis <= 0 Or dblMinorAxis <= 0 Then
    If Not TypeOf pGeometry.SpatialReference Is IGeographicCoordinateSystem Then
      MsgBox "Unexpected Spatial Reference:" & vbCrLf & _
          "The function 'FeaturePlanetOgraphicToPlanetOcentric' only accepts geometries with geographic projections." & vbCrLf & _
          "This geometry has spatial reference '" & pGeometry.SpatialReference.Name & "..."
      Set FeaturePlanetOGraphicToPlanetOCentric = Nothing
    End If
    Dim pGCS As IGeographicCoordinateSystem
    Set pGCS = pGeometry.SpatialReference
    
    Dim pEllipsoid As ISpheroid
    Set pEllipsoid = pGCS.Datum.Spheroid
    
  '  Dim dblFlattening As Double
    dblMajorAxis = pEllipsoid.SemiMajorAxis
    dblMinorAxis = pEllipsoid.SemiMinorAxis
  '  dblFlattening = pEllipsoid.Flattening
  End If
  
  Dim pPointColl As IPointCollection
  
  Dim pOutput As IPointCollection
  Dim pPoint As IPoint
  Dim pClone As IClone
  
  Dim dblLongitude As Double
  Dim dblLatitude As Double
  Dim dblNewLongitude As Double
  Dim dblNewLatitude As Double
  
  Dim pEnumVertex As IEnumVertex2
  Dim lngOutPart As Long
  Dim lngOutVertex As Long
  
  Dim lngIndex As Long
  
  ' IF A POINT, JUST CONVERT LATITUDE AND LONGITUDE AND RETURN NEW POINT
  ' IF POLYLINE, POLYGON OR MULTIPOINT, THEN CREATE NEW IPointCollection BY CLONING ORIGINAL SHAPE.  THEN JUST
  ' ADJUST EACH POINT IN POINT COLLECTION
  
  If TypeOf pGeometry Is IPoint Then
    Dim pNewPoint As IPoint
    Set pNewPoint = New Point
    Set pNewPoint.SpatialReference = pGeometry.SpatialReference
    
    Set pPoint = pGeometry
    dblLongitude = pPoint.x
    dblLatitude = pPoint.Y
    
    XYOGraphicToOCentric dblLongitude, dblLatitude, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
    pNewPoint.PutCoords dblNewLongitude, dblNewLatitude
        
    Set FeaturePlanetOGraphicToPlanetOCentric = pNewPoint
    
  ElseIf TypeOf pGeometry Is IPolyline Then
    
    Set pClone = pGeometry
    Dim pNewPolyline As IPolyline
    Set pNewPolyline = pClone.Clone
    Set pPointColl = pNewPolyline
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      XYOGraphicToOCentric pPoint.x, pPoint.Y, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
      pEnumVertex.put_Y dblNewLatitude
      pEnumVertex.put_X dblNewLongitude
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeaturePlanetOGraphicToPlanetOCentric = pNewPolyline
    
  ElseIf TypeOf pGeometry Is IPolygon Then
    
    Set pClone = pGeometry
    Dim pNewPolygon As IPolygon
    Set pNewPolygon = pClone.Clone
    
    Set pPointColl = pNewPolygon
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      XYOGraphicToOCentric pPoint.x, pPoint.Y, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
      pEnumVertex.put_Y dblNewLatitude
      pEnumVertex.put_X dblNewLongitude
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeaturePlanetOGraphicToPlanetOCentric = pNewPolygon
    
  ElseIf TypeOf pGeometry Is IMultipoint Then
    
    Set pClone = pGeometry
    Dim pNewMultipoint As IMultipoint
    Set pNewMultipoint = pClone.Clone
    Set pPointColl = pNewMultipoint
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      XYOGraphicToOCentric pPoint.x, pPoint.Y, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
      pEnumVertex.put_Y dblNewLatitude
      pEnumVertex.put_X dblNewLongitude
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeaturePlanetOGraphicToPlanetOCentric = pNewMultipoint
  
  Else
    MsgBox "Unexpected Geometry Type:" & vbCrLf & _
        "The function 'PlanetOCentricToPlanetOGraphic' only accepts points, polylines, polygons or multipoints."
    Set FeaturePlanetOGraphicToPlanetOCentric = Nothing
  End If

End Function


Public Function FeaturePlanetOCentricToPlanetOGraphic(pGeometry As IGeometry, Optional dblMajorAxis As Double = -999, _
    Optional dblMinorAxis As Double = -999, Optional dblLongShift As Double = 0) As IGeometry

  ' ASSUMES pGeometry IS IN A GEOGRAPHIC PROJECTION AND IS EITHER A POINT, POLYLINE, POLYGON OR MULTIPOINT
  
  If dblMajorAxis <= 0 Or dblMinorAxis <= 0 Then
    If Not TypeOf pGeometry.SpatialReference Is IGeographicCoordinateSystem Then
      MsgBox "Unexpected Spatial Reference:" & vbCrLf & _
          "The function 'FeaturePlanetOCentricToPlanetOGraphic' only accepts geometries with geographic projections." & vbCrLf & _
          "This geometry has spatial reference '" & pGeometry.SpatialReference.Name & "..."
      Set FeaturePlanetOCentricToPlanetOGraphic = Nothing
    End If
    Dim pGCS As IGeographicCoordinateSystem
    Set pGCS = pGeometry.SpatialReference
    
    Dim pEllipsoid As ISpheroid
    Set pEllipsoid = pGCS.Datum.Spheroid
    
  '  Dim dblFlattening As Double
    dblMajorAxis = pEllipsoid.SemiMajorAxis
    dblMinorAxis = pEllipsoid.SemiMinorAxis
  '  dblFlattening = pEllipsoid.Flattening
  End If
  
  Dim pPointColl As IPointCollection
  
  Dim pOutput As IPointCollection
  Dim pPoint As IPoint
  Dim pClone As IClone
  
  Dim dblLongitude As Double
  Dim dblLatitude As Double
  Dim dblNewLongitude As Double
  Dim dblNewLatitude As Double
  
  Dim pEnumVertex As IEnumVertex2
  Dim lngOutPart As Long
  Dim lngOutVertex As Long
  
  Dim lngIndex As Long
  
  ' IF A POINT, JUST CONVERT LATITUDE AND LONGITUDE AND RETURN NEW POINT
  ' IF POLYLINE, POLYGON OR MULTIPOINT, THEN CREATE NEW IPointCollection BY CLONING ORIGINAL SHAPE.  THEN JUST
  ' ADJUST EACH POINT IN POINT COLLECTION
  
  If TypeOf pGeometry Is IPoint Then
    Dim pNewPoint As IPoint
    Set pNewPoint = New Point
    Set pNewPoint.SpatialReference = pGeometry.SpatialReference
    
    Set pPoint = pGeometry
    dblLongitude = pPoint.x
    dblLatitude = pPoint.Y
    
    XYOCentricToOGraphic dblLongitude, dblLatitude, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
    pNewPoint.PutCoords dblNewLongitude, dblNewLatitude
        
    Set FeaturePlanetOCentricToPlanetOGraphic = pNewPoint
    
  ElseIf TypeOf pGeometry Is IPolyline Then
    
    Set pClone = pGeometry
    Dim pNewPolyline As IPolyline
    Set pNewPolyline = pClone.Clone
    Set pPointColl = pNewPolyline
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      XYOCentricToOGraphic pPoint.x, pPoint.Y, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
      pEnumVertex.put_Y dblNewLatitude
      pEnumVertex.put_X dblNewLongitude
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeaturePlanetOCentricToPlanetOGraphic = pNewPolyline
    
  ElseIf TypeOf pGeometry Is IPolygon Then
    
    Set pClone = pGeometry
    Dim pNewPolygon As IPolygon
    Set pNewPolygon = pClone.Clone
    
    Set pPointColl = pNewPolygon
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      XYOCentricToOGraphic pPoint.x, pPoint.Y, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
      pEnumVertex.put_Y dblNewLatitude
      pEnumVertex.put_X dblNewLongitude
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeaturePlanetOCentricToPlanetOGraphic = pNewPolygon
    
  ElseIf TypeOf pGeometry Is IMultipoint Then
    
    Set pClone = pGeometry
    Dim pNewMultipoint As IMultipoint
    Set pNewMultipoint = pClone.Clone
    Set pPointColl = pNewMultipoint
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      XYOCentricToOGraphic pPoint.x, pPoint.Y, dblMajorAxis, dblMinorAxis, dblLongShift, dblNewLongitude, dblNewLatitude
      pEnumVertex.put_Y dblNewLatitude
      pEnumVertex.put_X dblNewLongitude
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeaturePlanetOCentricToPlanetOGraphic = pNewMultipoint
  
  Else
    MsgBox "Unexpected Geometry Type:" & vbCrLf & _
        "The function 'PlanetOCentricToPlanetOGraphic' only accepts points, polylines, polygons or multipoints."
    Set FeaturePlanetOCentricToPlanetOGraphic = Nothing
  End If

End Function

Public Function FeatureLongitudeShift(pGeometry As IGeometry, dblLongShift As Double) As IGeometry
  
  Dim pPointColl As IPointCollection
  
  Dim pOutput As IPointCollection
  Dim pPoint As IPoint
  Dim pClone As IClone
  
  Dim dblLongitude As Double
  Dim dblLatitude As Double
  Dim dblNewLongitude As Double
  Dim dblNewLatitude As Double
  
  Dim pEnumVertex As IEnumVertex2
  Dim lngOutPart As Long
  Dim lngOutVertex As Long
  
  Dim lngIndex As Long
  
  ' IF A POINT, JUST CONVERT LATITUDE AND LONGITUDE AND RETURN NEW POINT
  ' IF POLYLINE, POLYGON OR MULTIPOINT, THEN CREATE NEW IPointCollection BY CLONING ORIGINAL SHAPE.  THEN JUST
  ' ADJUST EACH POINT IN POINT COLLECTION
  
  If TypeOf pGeometry Is IPoint Then
    Dim pNewPoint As IPoint
    Set pNewPoint = New Point
    Set pNewPoint.SpatialReference = pGeometry.SpatialReference
    
    Set pPoint = pGeometry
    dblLongitude = pPoint.x + dblLongShift
    dblLatitude = pPoint.Y
    
    pNewPoint.PutCoords dblLatitude, dblLongitude
        
    Set FeatureLongitudeShift = pNewPoint
    
  ElseIf TypeOf pGeometry Is IPolyline Then
    
    Set pClone = pGeometry
    Dim pNewPolyline As IPolyline
    Set pNewPolyline = pClone.Clone
    Set pPointColl = pNewPolyline
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      pEnumVertex.put_Y pPoint.Y
      pEnumVertex.put_X pPoint.x + dblLongShift
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeatureLongitudeShift = pNewPolyline
    
  ElseIf TypeOf pGeometry Is IPolygon Then
    
    Set pClone = pGeometry
    Dim pNewPolygon As IPolygon
    Set pNewPolygon = pClone.Clone
    
    Set pPointColl = pNewPolygon
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      pEnumVertex.put_Y pPoint.Y
      pEnumVertex.put_X pPoint.x + dblLongShift
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeatureLongitudeShift = pNewPolygon
    
  ElseIf TypeOf pGeometry Is IMultipoint Then
    
    Set pClone = pGeometry
    Dim pNewMultipoint As IMultipoint
    Set pNewMultipoint = pClone.Clone
    Set pPointColl = pNewMultipoint
    Set pEnumVertex = pPointColl.EnumVertices
    Set pPoint = New Point
    
    pEnumVertex.Reset
    pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Do While lngOutVertex > -1

      pEnumVertex.put_Y pPoint.Y
      pEnumVertex.put_X pPoint.x + dblLongShift
      
      pEnumVertex.QueryNext pPoint, lngOutPart, lngOutVertex
    
    Loop
    
    Set FeatureLongitudeShift = pNewMultipoint
  
  Else
    MsgBox "Unexpected Geometry Type:" & vbCrLf & _
        "The function 'FeatureLongitudeShift' only accepts points, polylines, polygons or multipoints."
    Set FeatureLongitudeShift = Nothing
  End If

End Function

Public Sub XYOCentricToOGraphic(dblLongitude As Double, dblLatitude As Double, dblMajorAxis As Double, dblMinorAxis As Double, _
    dblLongitudeShift As Double, dblNewLongitude As Double, dblNewLatitude As Double)
  
  ' ORIGINAL AVENUE CODE FROM View.Ocentric2Ographic
'  theLon = pt.GetX
'  theLon = theLon + theLonShift
'
'  theLat = pt.GetY
'  theLat = theLat * Number.GetPi / 180
'  theLat = ((((theMajorAx / theMinorAx)^2) * (theLat.tan))).atan
'  theLat = theLat * (180 / Number.GetPi)

  dblNewLongitude = dblLongitude + dblLongitudeShift
  dblNewLatitude = AsDegrees(Atn(((dblMajorAxis / dblMinorAxis) ^ 2) * (Tan(AsRadians(dblLatitude)))))

End Sub

Public Sub XYOGraphicToOCentric(dblLongitude As Double, dblLatitude As Double, dblMajorAxis As Double, dblMinorAxis As Double, _
    dblLongitudeShift As Double, dblNewLongitude As Double, dblNewLatitude As Double)
  
  ' ORIGINAL AVENUE CODE FROM View.Ographic2Ocentric
'  theLon = pt.GetX
'  theLon = theLon + theLonShift
'
'  theLat = pt.GetY
'  theLat = theLat * Number.GetPi / 180
'  theLat = (((theLat.tan) / ((theMajorAx / theMinorAx)^2))).atan
'  theLat = theLat * (180 / Number.GetPi)


  dblNewLongitude = dblLongitude + dblLongitudeShift
  dblNewLatitude = AsDegrees(Atn(Tan(AsRadians(dblLatitude)) / ((dblMajorAxis / dblMinorAxis) ^ 2)))

End Sub

Public Function WrapToBoundary(pGeometry As IGeometry, dblXMin As Double, dblXMax As Double, dblYMin As Double, dblYMax As Double, _
      Optional pMxDoc As IMxDocument) As IGeometry
  
  Dim dblXRange As Double
  Dim dblYRange As Double
  dblXRange = dblXMax - dblXMin
  dblYRange = dblYMax - dblYMin
  
  Dim dblTestX As Double
  Dim dblTestY As Double
  Dim pPoint As IPoint
  Dim pNewPoint As IPoint
  
  Dim pTopoOp As ITopologicalOperator3
  Dim pSpRef As ISpatialReference
  Set pSpRef = pGeometry.SpatialReference
    
  Dim lngIndex As Long
  
  If TypeOf pGeometry Is IPoint Then
    
    Set pPoint = pGeometry
    dblTestX = pPoint.x
    dblTestY = pPoint.Y
    
    Do Until dblTestX <= dblXMax
      dblTestX = dblTestX - dblXRange
    Loop
    Do Until dblTestX >= dblXMin
      dblTestX = dblTestX + dblXRange
    Loop
    
    Do Until dblTestY <= dblYMax
      dblTestY = dblTestY - dblYRange
    Loop
    Do Until dblTestY >= dblYMin
      dblTestY = dblTestY + dblYRange
    Loop
    
    Set pNewPoint = New Point
    pNewPoint.PutCoords dblTestX, dblTestY
    Set pNewPoint.SpatialReference = pSpRef
    
    Set WrapToBoundary = pNewPoint
  Else
    
    ' START BY MAKING A SET OF CLIPPING ENVELOPES COVERING THE EXTENT OF THE SHAPE
    Dim pEnvelope As IEnvelope
    Set pEnvelope = pGeometry.Envelope
    
    Dim dblEnvMaxX As Double
    Dim dblEnvMinX As Double
    Dim dblEnvMaxY As Double
    Dim dblEnvMinY As Double
    
    dblEnvMaxX = pEnvelope.XMax
    dblEnvMinX = pEnvelope.XMin
    dblEnvMaxY = pEnvelope.YMax
    dblEnvMinY = pEnvelope.YMin
    
    Dim dblRunningX As Double
    Dim dblRunningY As Double
    
    dblRunningX = dblXMin
    dblRunningY = dblYMin
    
    Dim dblShiftX As Double
    Dim dblShiftY As Double
    Dim dblShiftBaseX As Double
    Dim dblShiftBaseY As Double
    
    dblShiftBaseX = 0
    dblShiftBaseY = 0
    
    Dim dblRunningBaseY As Double
    
    ' IDENTIFY BOTTOM LEFT CORNER
    Do Until dblRunningX <= dblEnvMinX
      dblRunningX = dblRunningX - dblXRange
      dblShiftBaseX = dblShiftBaseX - dblXRange
    Loop
    Do Until dblRunningY <= dblEnvMinY
      dblRunningY = dblRunningY - dblYRange
      dblShiftBaseY = dblShiftBaseY - dblYRange
    Loop
    dblRunningBaseY = dblRunningY
    
    ' MAKE SET OF RECTANGULAR POLYGONS, WHERE EACH RECTANGLE IS EQUAL TO SIZE OF BOUNDARY, AND SET ENTIRELY COVERS SHAPE
    Dim pRectArray As esriSystem.IArray
    Set pRectArray = New esriSystem.Array
    
    Dim pShiftArray As esriSystem.IArray
    Set pShiftArray = New esriSystem.Array
    Dim pShiftSubArray As esriSystem.IDoubleArray
    
    Dim pRect As IPointCollection
    
    Dim pRelOp As IRelationalOperator
    Set pRelOp = pGeometry
    
    dblShiftX = dblShiftBaseX
    dblShiftY = dblShiftBaseY
    
    Do Until dblRunningX >= dblEnvMaxX
      Do Until dblRunningY >= dblEnvMaxY
        Set pRect = New Polygon
        
        Set pNewPoint = New Point
        pNewPoint.PutCoords dblRunningX, dblRunningY
        pRect.AddPoint pNewPoint
        
        Set pNewPoint = New Point
        pNewPoint.PutCoords dblRunningX, dblRunningY + dblYRange
        pRect.AddPoint pNewPoint
        
        Set pNewPoint = New Point
        pNewPoint.PutCoords dblRunningX + dblXRange, dblRunningY + dblYRange
        pRect.AddPoint pNewPoint
        
        Set pNewPoint = New Point
        pNewPoint.PutCoords dblRunningX + dblXRange, dblRunningY
        pRect.AddPoint pNewPoint
        
        Set pNewPoint = New Point
        pNewPoint.PutCoords dblRunningX, dblRunningY
        pRect.AddPoint pNewPoint
        
        Set pTopoOp = pRect
        pTopoOp.Simplify
        
        ' ONLY CONSIDER RECTANGLES THAT ACTUALLY INTERSECT THE ORIGINAL SHAPE
        If Not pRelOp.Disjoint(pRect) Then
          pRectArray.Add pRect
          Set pShiftSubArray = New esriSystem.DoubleArray
          pShiftSubArray.Add dblShiftX
          pShiftSubArray.Add dblShiftY
          pShiftArray.Add pShiftSubArray
          
        End If
        dblRunningY = dblRunningY + dblYRange
        dblShiftY = dblShiftY + dblYRange
      Loop
      dblRunningY = dblRunningBaseY
      dblShiftY = dblShiftBaseY
      dblRunningX = dblRunningX + dblXRange
      dblShiftX = dblShiftX + dblXRange
    Loop
    
    Dim pOutputColl As IGeometryCollection
    Set pOutputColl = New GeometryBag
    Dim pUnionTopoOp As ITopologicalOperator
    
    Dim pRectPoly As IPolygon
    Dim pClipPointColl As IPointCollection
    
    Dim pClone As IClone
    
    If TypeOf pGeometry Is IPolyline Then
      Dim pPolyline As IPolyline
      Set pPolyline = pGeometry
      Set pUnionTopoOp = New Polyline
      
      Dim pClipPolyline As IPolyline
      For lngIndex = 0 To pRectArray.Count - 1
        Set pRectPoly = pRectArray.Element(lngIndex)
        Set pClone = pPolyline
        Set pClipPolyline = pClone.Clone
        Set pTopoOp = pClipPolyline
        pTopoOp.Simplify
        
        pTopoOp.Clip pRectPoly.Envelope
        Set pClipPointColl = pClipPolyline
        
        ShiftPointsToWrapBoundary pClipPointColl, pShiftArray.Element(lngIndex)
        
        Set pTopoOp = pClipPolyline
        pTopoOp.Simplify
        
        pOutputColl.AddGeometry pClipPolyline
      Next lngIndex
      
    ElseIf TypeOf pGeometry Is IPolygon Then
      Dim pPolygon As IPolygon
      Set pPolygon = pGeometry
      Set pUnionTopoOp = New Polygon
      
      Dim pClipPolygon As IPolygon
      For lngIndex = 0 To pRectArray.Count - 1
        Set pRectPoly = pRectArray.Element(lngIndex)
        Set pClone = pPolygon
        Set pClipPolygon = pClone.Clone
        Set pTopoOp = pClipPolygon
        pTopoOp.Simplify
        pTopoOp.ClipDense pRectPoly.Envelope, 0.1
        Set pClipPointColl = pClipPolygon
        
        ShiftPointsToWrapBoundary pClipPointColl, pShiftArray.Element(lngIndex)
        Set pTopoOp = pClipPolygon
        pTopoOp.Simplify
        
        pOutputColl.AddGeometry pClipPolygon
      Next lngIndex
    
    ElseIf TypeOf pGeometry Is IMultipoint Then
      Dim pMultipoint As IMultipoint
      Set pMultipoint = pGeometry
      Set pUnionTopoOp = New Multipoint
      
      Dim pClipMultipoint As IMultipoint
      For lngIndex = 0 To pRectArray.Count - 1
        Set pRectPoly = pRectArray.Element(lngIndex)
        Set pClone = pMultipoint
        Set pClipMultipoint = pClone.Clone
        Set pTopoOp = pClipMultipoint
        pTopoOp.Clip pRectPoly.Envelope
        Set pClipPointColl = pClipMultipoint
        
        ShiftPointsToWrapBoundary pClipPointColl, pShiftArray.Element(lngIndex)
        Set pTopoOp = pClipMultipoint
        pTopoOp.Simplify
        
        pOutputColl.AddGeometry pClipMultipoint
      Next lngIndex
      
    End If
    pUnionTopoOp.ConstructUnion pOutputColl
    pUnionTopoOp.Simplify
    Set WrapToBoundary = pUnionTopoOp
  End If

End Function

Private Sub ShiftPointsToWrapBoundary(pPointColl As IPointCollection, pShiftSubArray As esriSystem.IDoubleArray)
  
  ' THIS FUNCTION ASSUMES THAT SHAPE HAS ALREADY BEEN CLIPPED TO RECTANGLE REPRESENTING WRAP BOUNDARY INCREMENTS
  
  Dim pTransform As ITransform2D
  Set pTransform = pPointColl
  pTransform.Move -pShiftSubArray.Element(0), -pShiftSubArray.Element(1)

End Sub

Public Function HSin(dblRadians As Double) As Double

  HSin = (Exp(dblRadians) - Exp(-dblRadians)) / 2

End Function

Public Function HCos(dblRadians As Double) As Double

  HCos = (Exp(dblRadians) + Exp(-dblRadians)) / 2

End Function

Public Function HTan(dblRadians As Double) As Double

  HTan = (Exp(dblRadians) - Exp(-dblRadians)) / (Exp(dblRadians) + Exp(-dblRadians))

End Function

Public Function HArcSin(dblValue As Double) As Double

  HArcSin = Log(dblValue + Sqr(dblValue * dblValue + 1))

End Function

Public Function HArcCos(dblValue As Double) As Double

  HArcCos = Log(dblValue + Sqr(dblValue * dblValue - 1))

End Function

Public Function HArcTan(dblValue As Double) As Double

  HArcTan = Log((1 + dblValue) / (1 - dblValue)) / 2

End Function

Public Function CalcInternalAngle(theP As IPoint, theQ As IPoint, theR As IPoint, Optional dblAngleDev As Double) As Double
 
On Error GoTo err
' CalcCheckClockwise
' Jenness Enterprises <www.jennessent.com)>
' Given 3 consecutive points, this scripts calculates the internal angle

' INTERNAL ANGLE WITH LAW OF COSINES;
'       c^2 = a^2 + b^2 - (2ab * Cos C), OR
'       Cos C = (a^2 +b^2 - c^2)/(2ab)

  Dim dblLenPQ As Double
  Dim dblLenQR As Double
  Dim dblLenRP As Double
  
  dblLenPQ = ((theP.x - theQ.x) ^ 2 + (theP.Y - theQ.Y) ^ 2) ^ (0.5)
  dblLenQR = ((theQ.x - theR.x) ^ 2 + (theQ.Y - theR.Y) ^ 2) ^ (0.5)
  dblLenRP = ((theR.x - theP.x) ^ 2 + (theR.Y - theP.Y) ^ 2) ^ (0.5)
  
  CalcInternalAngle = (((dblLenPQ ^ 2) + (dblLenQR ^ 2) - (dblLenRP ^ 2)) / (2 * dblLenPQ * dblLenQR))
  CalcInternalAngle = ArcCosJen(CalcInternalAngle)
  CalcInternalAngle = Round(RadToDeg(CalcInternalAngle), 10)
  dblAngleDev = 180 - CalcInternalAngle

    Exit Function
err:
  Dim dblbearing1 As Double
  Dim dblBearing2 As Double
  dblbearing1 = Round(CalcBearing(theP, theQ))
  dblBearing2 = Round(CalcBearing(theQ, theR))
  If (dblbearing1 = dblBearing2) Then
    CalcInternalAngle = 180
    dblAngleDev = 0
  Else
    CalcInternalAngle = 0
    dblAngleDev = 180
  End If

End Function

Public Function CreateCircleAroundPoint(pOrigin As IPoint, dblRadius As Double, lngPtCount As Long)

  Dim dblInterval As Double
  dblInterval = 360 / lngPtCount
  Dim dblIndex As Double
  
  Dim pCircle As IPointCollection
  Set pCircle = New Polygon
  Dim pGeom As IGeometry
  Set pGeom = pCircle
  Set pGeom.SpatialReference = pOrigin.SpatialReference
  
  Dim pNewPoint As IPoint
  Dim pTopoOp As ITopologicalOperator
  
  dblIndex = 0
  Do Until dblIndex >= 360
    
    CalcPointLine pOrigin, dblRadius, dblIndex, pNewPoint
    pCircle.AddPoint pNewPoint
    
    dblIndex = dblIndex + dblInterval
  Loop
  
  Dim pFinal As IPolygon
  Set pFinal = pCircle
  pFinal.Close
  Set pTopoOp = pFinal
  pTopoOp.Simplify
  
  Set CreateCircleAroundPoint = pFinal

End Function


Public Function CreateBoxAroundPoint(pOrigin As IPoint, dblXDistFromOrigin As Double, dblYDistFromOrigin As Double) As IPolygon
  
  Dim pBox As IPointCollection
  Set pBox = New Polygon
  Dim pGeom As IGeometry
  Set pGeom = pBox
  Set pGeom.SpatialReference = pOrigin.SpatialReference
  
  Dim pNewPoint As IPoint
  Dim pTopoOp As ITopologicalOperator
  
  Set pNewPoint = New Point
  pNewPoint.PutCoords pOrigin.x - dblXDistFromOrigin, pOrigin.Y - dblYDistFromOrigin
  pBox.AddPoint pNewPoint
  
  Set pNewPoint = New Point
  pNewPoint.PutCoords pOrigin.x - dblXDistFromOrigin, pOrigin.Y + dblYDistFromOrigin
  pBox.AddPoint pNewPoint
  
  Set pNewPoint = New Point
  pNewPoint.PutCoords pOrigin.x + dblXDistFromOrigin, pOrigin.Y + dblYDistFromOrigin
  pBox.AddPoint pNewPoint
  
  Set pNewPoint = New Point
  pNewPoint.PutCoords pOrigin.x + dblXDistFromOrigin, pOrigin.Y - dblYDistFromOrigin
  pBox.AddPoint pNewPoint
  
  Dim pFinal As IPolygon
  Set pFinal = pBox
  pFinal.Close
  Set pTopoOp = pFinal
  pTopoOp.Simplify
  
  Set CreateBoxAroundPoint = pFinal

End Function

Public Function EstimateDistanceOnSphere(pGeom As IGeometry, dblMeters As Double, _
      Optional booFoundProblems As Boolean = False, Optional strProblemReason As String, Optional dblAz As Double = 45) As Double
  
  If pGeom Is Nothing Then
    EstimateDistanceOnSphere = dblMeters
    booFoundProblems = True
    strProblemReason = "Empty Geometry"
    Exit Function
  End If
  
  Dim pSpRef As ISpatialReference
  Set pSpRef = pGeom.SpatialReference
  Dim pPrjCS As IProjectedCoordinateSystem
  Dim pGeoCS As IGeographicCoordinateSystem
  booFoundProblems = False
  
  If pSpRef Is Nothing Then
    EstimateDistanceOnSphere = dblMeters
    booFoundProblems = True
    strProblemReason = "Spatial Reference Missing"
    Exit Function
  End If
  If TypeOf pGeom.SpatialReference Is IUnknownCoordinateSystem Then
    EstimateDistanceOnSphere = dblMeters
    booFoundProblems = True
    strProblemReason = "Spatial Reference Unknown"
    Exit Function
  End If
  If TypeOf pGeom.SpatialReference Is IProjectedCoordinateSystem Then
    EstimateDistanceOnSphere = dblMeters
    booFoundProblems = True
    strProblemReason = "Spatial Reference Projected"
    Exit Function
  End If
  
  Set pGeoCS = pSpRef
  
  Dim pPoint As IPoint
  If TypeOf pGeom Is IPoint Then
    Set pPoint = pGeom
  Else
    Dim pEnv As IEnvelope
    Set pEnv = pGeom.Envelope
    Set pPoint = New Point
    pPoint.PutCoords (pEnv.XMax - pEnv.XMin) / 2 + pEnv.XMin, (pEnv.YMax - pEnv.YMin) / 2 + pEnv.YMin
  End If
  
  Dim pSpheroid As ISpheroid
  Dim pDatum As IDatum
  Set pDatum = pGeoCS.Datum
  Set pSpheroid = pDatum.Spheroid
  
  Dim pNewPoint As IPoint
  Set pNewPoint = New Point
  Dim dblAZ2 As Double
  PointLineVincentyPerPoint2 pPoint, dblMeters, dblAz, pNewPoint, dblAZ2, pSpheroid.SemiMajorAxis, pSpheroid.SemiMinorAxis
  
  EstimateDistanceOnSphere = (((pPoint.x - pNewPoint.x) ^ 2) + ((pPoint.Y - pNewPoint.Y) ^ 2)) ^ (0.5)

  Set pSpRef = Nothing
  Set pPrjCS = Nothing
  Set pGeoCS = Nothing
  Set pPoint = Nothing
  Set pEnv = Nothing
  Set pSpheroid = Nothing
  Set pDatum = Nothing
  Set pNewPoint = Nothing

End Function


Public Function BufferGeographic(ByVal pOrigGeom As IGeometry, dblMeters As Double, _
      Optional booFoundProblems As Boolean = False, Optional strProblemReason As String) As IPolygon

  Dim pClone As IClone
  Dim pGeom As IGeometry
  Set pClone = pOrigGeom
  Set pGeom = pClone.Clone
  Dim pTempEnv As IEnvelope
  
  If TypeOf pGeom Is IEnvelope Then
    Set pTempEnv = pGeom
    Set pGeom = EnvelopeToPolygon(pTempEnv)
    Set pTempEnv = Nothing
  End If
  
  Dim pOrigSpRef As ISpatialReference
  Set pOrigSpRef = pGeom.SpatialReference
  If Not TypeOf pOrigSpRef Is IGeographicCoordinateSystem Then
    booFoundProblems = True
    strProblemReason = "Spatial Reference Not Geographic"
    Exit Function
  End If
  
  Dim dblXMin As Double
  Dim dblXMax As Double
  Dim dblYMin As Double
  Dim dblYMax As Double
  
  Dim pPoint As IPoint
  Dim pEnv As IEnvelope
  
  If TypeOf pGeom Is IPoint Then
    Set pPoint = pGeom
    dblXMin = pPoint.x - (dblMeters * 10)
    dblXMax = pPoint.x + (dblMeters * 10)
    dblYMin = pPoint.Y - (dblMeters * 10)
    dblYMax = pPoint.Y + (dblMeters * 10)
  Else
    Set pPoint = New Point
    Set pEnv = pGeom.Envelope
    pPoint.PutCoords (pEnv.XMax - pEnv.XMin) / 2 + pEnv.XMin, (pEnv.YMax - pEnv.YMin) / 2 + pEnv.YMin
    dblXMin = pEnv.XMin - (dblMeters * 10)
    dblXMax = pEnv.XMax + (dblMeters * 10)
    dblYMin = pEnv.YMin - (dblMeters * 10)
    dblYMax = pEnv.YMax + (dblMeters * 10)
  End If
    
  ' PROJECT INTO AZIMUTHAL EQUIDISTANT
  Dim pSpRefFact As ISpatialReferenceFactory3
  Set pSpRefFact = New SpatialReferenceEnvironment
  Dim pPrjCS As IProjectedCoordinateSystem3
  Set pPrjCS = pSpRefFact.CreateProjectedCoordinateSystem(esriSRProjCS_World_AzimuthalEquidistant)
  Dim pSpRef As ISpatialReference
  Set pSpRef = pPrjCS
  pPrjCS.CentralMeridian(True) = pPoint.x
  pPrjCS.LatitudeOfOrigin = pPoint.Y
  
  If Not MyGeomCheckSpRefDomain(pSpRef) Then
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
  End If
  
  pGeom.Project pSpRef
  Dim pTopoOp As ITopologicalOperator
  Set pTopoOp = pGeom
  pTopoOp.Simplify
  
  Dim pBuffer As IPolygon
  Set pBuffer = pTopoOp.Buffer(dblMeters)
  Set pTopoOp = pBuffer
  pTopoOp.Simplify
  
  pBuffer.Project pOrigSpRef
  pTopoOp.Simplify
  
  Set BufferGeographic = pBuffer
  

  Set pClone = Nothing
  Set pGeom = Nothing
  Set pOrigSpRef = Nothing
  Set pPoint = Nothing
  Set pEnv = Nothing
  Set pSpRefFact = Nothing
  Set pPrjCS = Nothing
  Set pSpRef = Nothing
  Set pSpRefRes = Nothing
  Set pTopoOp = Nothing
  Set pBuffer = Nothing
  Set pTempEnv = Nothing

End Function

Private Function MyGeomCheckSpRefDomain(pSpRef As ISpatialReference) As Boolean
  On Error GoTo ErrHand

  
  Dim dXmax As Double
  Dim dYmax As Double
  Dim dXmin As Double
  Dim dYmin As Double
  
  'get the xy domain extent of the dataset
  
  pSpRef.GetDomain dXmin, dXmax, dYmin, dYmax
  MyGeomCheckSpRefDomain = True
  
ErrHand:
  MyGeomCheckSpRefDomain = False

End Function


Public Function CreateEllipseAroundPoint(pCenter As IPoint, dblSemiMajor As Double, _
  dblSemiMinor As Double, dblFlatOrientationCCWFromHorizontal As Double, Optional lngNumPoints As Long = 360) As IPolygon

  Set CreateEllipseAroundPoint = Nothing
  
  Dim pPointColl As IPointCollection
  Set pPointColl = New Polygon
  
  Dim dblAngleInterval As Double
  dblAngleInterval = 360 / lngNumPoints
  
  Dim dblX As Double
  Dim dblY As Double
  Dim pNewPoint As IPoint
  
  Dim dblCenterX As Double
  Dim dblCenterY As Double
  dblCenterX = pCenter.x
  dblCenterY = pCenter.Y
  
  Dim dblAngle As Double
  Dim dblRadians As Double
  Dim dblRadiansFromNorth As Double
  dblRadiansFromNorth = dblPI * (dblFlatOrientationCCWFromHorizontal / 180)
  
  For dblAngle = 0 To 360 Step dblAngleInterval
    dblRadians = dblPI * (dblAngle / 180)
    dblX = dblCenterX + (dblSemiMajor * Cos(dblRadians) * Cos(dblRadiansFromNorth)) - _
        (dblSemiMinor * Sin(dblRadians) * Sin(dblRadiansFromNorth))
    dblY = dblCenterY + (dblSemiMajor * Cos(dblRadians) * Sin(dblRadiansFromNorth)) + _
        (dblSemiMinor * Sin(dblRadians) * Cos(dblRadiansFromNorth))
    Set pNewPoint = New Point
    pNewPoint.PutCoords dblX, dblY
    
    pPointColl.AddPoint pNewPoint
    
  Next dblAngle
  
  Dim pEllipse As IPolygon
  Set pEllipse = pPointColl
  pEllipse.Close
    
  Set pEllipse.SpatialReference = pCenter.SpatialReference
  
  Dim pTopoOp As ITopologicalOperator
  Set pTopoOp = pEllipse
  pTopoOp.Simplify
  
  Set CreateEllipseAroundPoint = pEllipse

End Function




Public Function CreateCrossAroundPoint(pCenter As IPoint, dblVerticalHalfLength As Double, dblHorizontalHalfLength As Double) As IPolyline

  Dim pSegColl As ISegmentCollection
  Set pSegColl = New Polyline
  
  Dim pToPoint As IPoint
  
  Dim pLine As ILine
  
  ' NORTH
  Set pToPoint = New Point
  pToPoint.PutCoords pCenter.x, pCenter.Y + dblVerticalHalfLength
  Set pLine = New esriGeometry.Line
  pLine.FromPoint = pCenter
  pLine.ToPoint = pToPoint
  pSegColl.AddSegment pLine
  
  ' EAST
  Set pToPoint = New Point
  pToPoint.PutCoords pCenter.x + dblHorizontalHalfLength, pCenter.Y
  Set pLine = New esriGeometry.Line
  pLine.FromPoint = pCenter
  pLine.ToPoint = pToPoint
  pSegColl.AddSegment pLine
  
  ' SOUTH
  Set pToPoint = New Point
  pToPoint.PutCoords pCenter.x, pCenter.Y - dblVerticalHalfLength
  Set pLine = New esriGeometry.Line
  pLine.FromPoint = pCenter
  pLine.ToPoint = pToPoint
  pSegColl.AddSegment pLine
  
  ' WEST
  Set pToPoint = New Point
  pToPoint.PutCoords pCenter.x - dblHorizontalHalfLength, pCenter.Y
  Set pLine = New esriGeometry.Line
  pLine.FromPoint = pCenter
  pLine.ToPoint = pToPoint
  pSegColl.AddSegment pLine
  
  Dim pPolyline As IPolyline
  Set pPolyline = pSegColl
  Set pPolyline.SpatialReference = pCenter.SpatialReference
  
  Dim pTopoOp As ITopologicalOperator
  Set pTopoOp = pPolyline
  pTopoOp.Simplify
  
  Set CreateCrossAroundPoint = pPolyline

End Function

Public Function ModDouble(dblValue As Double, dblModValue As Double) As Double
  

'    ModDouble = dblValue - (dblModValue * CLng(dblValue / dblModValue))

    Dim dblVal As Double
    Dim dblRatio As Double
    dblRatio = dblValue / dblModValue
    dblVal = Round(dblRatio)
    If dblVal > dblRatio Then dblVal = dblVal - 1
    ModDouble = dblValue - (dblModValue * dblVal)
    
End Function
Public Function ReturnMeanDir(dblCompassDirs() As Double) As Double

  Dim dblSumC As Double
  Dim dblSumS As Double
  Dim lngIndex As Long
  Dim dblRadians As Double
  
  For lngIndex = 0 To UBound(dblCompassDirs)
    
    dblRadians = AsRadians(dblCompassDirs(lngIndex))
    dblSumC = dblSumC + Cos(dblRadians)
    dblSumS = dblSumS + Sin(dblRadians)
        
  Next lngIndex
  
  Dim dblR As Double
  dblR = Sqr(dblSumC ^ 2 + dblSumS ^ 2)
  
  Dim dblMeanDir As Double
  dblMeanDir = atan2(dblSumS, dblSumC)
  dblMeanDir = AsDegrees(dblMeanDir)
  If dblMeanDir < 0 Then
    dblMeanDir = dblMeanDir + 360
  End If
  
  ReturnMeanDir = dblMeanDir

End Function


Public Function ConvertDDtoDMS(dblVal As Double, lngDegrees As Long, lngMin As Long, dblSec As Double) As Long
  
  ConvertDDtoDMS = -999
  Dim dblTemp As Double
  lngDegrees = Fix(dblVal)
  dblTemp = Abs(dblVal - CDbl(lngDegrees)) * 60
  lngMin = Fix(dblTemp)
  dblSec = Abs(dblTemp - CDbl(lngMin)) * 60
  ConvertDDtoDMS = 1

End Function

Public Function ConvertDMStoDD(lngDegrees As Long, lngMin As Long, dblSec As Double) As Double

  If Sgn(lngDegrees) = -1 Then
    ConvertDMStoDD = CDbl(lngDegrees) - (CDbl(lngMin) / 60) - (dblSec / 3600)
  Else
    ConvertDMStoDD = CDbl(lngDegrees) + (CDbl(lngMin) / 60) + (dblSec / 3600)
  End If

End Function

Public Function CreateCircleAroundPointGeographic(pOrigin As IPoint, dblRadius As Double, lngPtCount As Long) As IPolygon
  
  Set CreateCircleAroundPointGeographic = Nothing
  
  Dim dblInterval As Double
  dblInterval = 360 / lngPtCount
  Dim dblIndex As Double
  
  Dim pCircle As IPointCollection
  Set pCircle = New Polygon
  Dim pGeom As IGeometry
  Set pGeom = pCircle
  Set pGeom.SpatialReference = pOrigin.SpatialReference
  
  Dim pNewPoint As IPoint
  Dim pTopoOp As ITopologicalOperator
  Dim dblAZ2 As Double
  
  dblIndex = 0
  
  Do Until dblIndex >= 360
    
    Set pNewPoint = New Point
    PointLineVincentyPerPoint2 pOrigin, dblRadius, dblIndex, pNewPoint, dblAZ2
    pCircle.AddPoint pNewPoint
    
    dblIndex = dblIndex + dblInterval
  Loop
  
  Dim pFinal As IPolygon
  Set pFinal = pCircle
  pFinal.Close
  Set pTopoOp = pFinal
  pTopoOp.Simplify
  
  Set CreateCircleAroundPointGeographic = pFinal

End Function

Public Function UnionGeometries(pGeomArray As esriSystem.IVariantArray) As IGeometry
  
'  Dim pMxDox As IMxDocument
'  Set pMxDoc = ThisDocument
  
  Dim pTopoOp As ITopologicalOperator
  Dim pGeom As IGeometry
  Dim pGeometryCollection As IGeometryCollection
  
  Set pGeometryCollection = New GeometryBag
  
  Set pGeom = pGeomArray.Element(0)
  Dim pSpRef As ISpatialReference
  Set pSpRef = pGeom.SpatialReference
  
  Dim lngGeomType As esriGeometryType
  lngGeomType = pGeom.GeometryType
  
  Dim lngIndex As Long
  For lngIndex = 0 To pGeomArray.Count - 1
    Set pGeom = pGeomArray.Element(lngIndex)
    
    
    If Not pGeom.IsEmpty Then
      pGeometryCollection.AddGeometry pGeom
    End If
  Next lngIndex
  
  Dim pNewGeom As IGeometry
  If lngGeomType = esriGeometryPoint Then
    Set pNewGeom = New Multipoint
  ElseIf lngGeomType = esriGeometryMultipoint Then
    Set pNewGeom = New Multipoint
  ElseIf lngGeomType = esriGeometryPolyline Then
    Set pNewGeom = New Polyline
  ElseIf lngGeomType = esriGeometryPolygon Then
    Set pNewGeom = New Polygon
  End If
  
  Set pTopoOp = pNewGeom
  pTopoOp.ConstructUnion pGeometryCollection
  pTopoOp.Simplify
  
  Set pNewGeom.SpatialReference = pSpRef
  
  Set UnionGeometries = pNewGeom
    
  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing

End Function

Public Function DistancePythagoreanNumbers(dblX1 As Double, dblY1 As Double, dblX2 As Double, dblY2 As Double) As Double

  DistancePythagoreanNumbers = Sqr(((dblX1 - dblX2) ^ 2) + ((dblY1 - dblY2) ^ 2))

End Function


Public Function ReturnVerticesAsDoubleArray(pGeometry As IGeometry) As Double()

  ' RETURNS A 2-DIMENSIONAL ARRAY OF X- AND Y-COORDINATES OF ALL POINTS
        
  Dim pPtColl As IPointCollection
  Set pPtColl = pGeometry
  Dim pTestPoint1 As IPoint
  Dim lngIndex1 As Long
  Dim lngPointCount As Long
  
  Set pTestPoint1 = New Point
  
  Dim dblReturn() As Double
  ReDim dblReturn(1, pPtColl.PointCount - 1)
  
  
  lngPointCount = pPtColl.PointCount
  
  If lngPointCount > 1 Then
    For lngIndex1 = 0 To lngPointCount - 1
      pPtColl.QueryPoint lngIndex1, pTestPoint1
      dblReturn(0, lngIndex1) = pTestPoint1.x
      dblReturn(1, lngIndex1) = pTestPoint1.Y
    Next lngIndex1
  Else
    pPtColl.QueryPoint 0, pTestPoint1
    dblReturn(0, 0) = pTestPoint1.x
    dblReturn(1, 0) = pTestPoint1.Y
  End If

  
  ReturnVerticesAsDoubleArray = dblReturn

End Function

Public Function CalcCheckClockwiseNumbers(dblPX As Double, dblPY As Double, dblQX As Double, _
    dblQY As Double, dblRX As Double, dblRY As Double, Optional dblDistance As Double) As Boolean
 
  ' CalcCheckClockwise
  ' Jenness Enterprises <www.jennessent.com)>
  ' Given 3 consecutive points, this scripts calculates whether the third point lies to the right
  ' (clockwise) or to the left (counter-clockwise) of the line connecting the first point to
  ' the second point.
  
  ' CLOCKWISE IF TRUE
  dblDistance = (dblQX * (dblRY - dblPY)) + (dblQY * (dblPX - dblRX)) - ((dblPX) * (dblRY)) _
        + ((dblPY) * (dblRX))
        
  CalcCheckClockwiseNumbers = dblDistance < 0

End Function

Public Function DistancePointToInfiniteLine(dblSegmentStartX As Double, dblSegmentStartY As Double, dblSegmentEndX As Double, _
    dblSegmentEndY As Double, dblPointX As Double, dblPointY As Double, Optional lngClockwise As JenClockwiseConstants) As Double
 
  ' DistancePointToInfiniteLine
  ' Jenness Enterprises <www.jennessent.com)>
  ' WILL CRASH IF SEGMENT START POINT COORDINATES ARE EQUAL TO SEGMENT END POINT COORDINATES
  ' Given 2 consecutive points defining a line with direction, this scripts calculates whether the third point lies to the right
  ' (clockwise) or to the left (counter-clockwise) of the line connecting the first point to the second point, and the distance
  ' from the point to the line.
  
  ' ASSUMES COORDINATES ARE PROJECTED!!!
  
  DistancePointToInfiniteLine = (((dblSegmentEndX - dblSegmentStartX) * (dblSegmentStartY - dblPointY)) - _
                 ((dblSegmentStartX - dblPointX) * (dblSegmentEndY - dblSegmentStartY))) / _
                 ((((dblSegmentEndX - dblSegmentStartX) ^ 2) + ((dblSegmentEndY - dblSegmentStartY) ^ 2)) ^ (0.5))
  
  If DistancePointToInfiniteLine < 0 Then
      lngClockwise = ENUM_CounterClockwise
  ElseIf DistancePointToInfiniteLine = 0 Then
      lngClockwise = Enum_OnLine
  Else
      lngClockwise = Enum_Clockwise
  End If
  
  
  DistancePointToInfiniteLine = Abs(DistancePointToInfiniteLine)

End Function

Public Function DistancePointToSegment(dblSegmentStartX As Double, dblSegmentStartY As Double, dblSegmentEndX As Double, _
    dblSegmentEndY As Double, dblPointX As Double, dblPointY As Double, Optional lngClockwise As JenClockwiseConstants, _
    Optional dblX_On_Segment As Double, Optional dblY_On_Segment As Double, Optional dblDistToInfiniteLine As Double) As Double
 
  ' DistancePointToSegment
  ' Jenness Enterprises <www.jennessent.com)>
  ' adapted from http://forums.codeguru.com/showthread.php?t=194400
  ' WILL CRASH IF SEGMENT START POINT COORDINATES ARE EQUAL TO SEGMENT END POINT COORDINATES
  ' Given 2 consecutive points defining a segment with direction, this scripts calculates whether the third point lies to the right
  ' (clockwise) or to the left (counter-clockwise) of the line connecting the first point to the second point, and the distance
  ' from the point to the segment.
  '
  ' ASSUMES COORDINATES ARE PROJECTED!!!
    
  Dim dblProportionAlongLine As Double
  ' values interpreted as follows:
  ' P is projection of 3rd point onto line
  ' dblProportionAlongLine = 0:  P = segment start point
  ' dblProportionAlongLine = 1:  P = segment end point
  ' dblProportionAlongLine < 0:  P is behind segment start point
  ' dblProportionAlongLine > 1:  P is beyond segment end point
  ' dblProportionAlongLine between 0 and 1:  P is between segment start and end points
  
  Dim dblNumerator As Double
  Dim dblDenom As Double
  dblNumerator = ((dblPointX - dblSegmentStartX) * (dblSegmentEndX - dblSegmentStartX)) + _
      ((dblPointY - dblSegmentStartY) * (dblSegmentEndY - dblSegmentStartY))
  dblDenom = ((dblSegmentEndX - dblSegmentStartX) * (dblSegmentEndX - dblSegmentStartX)) + _
      ((dblSegmentEndY - dblSegmentStartY) * (dblSegmentEndY - dblSegmentStartY))
      
  dblProportionAlongLine = dblNumerator / dblDenom
  
  dblX_On_Segment = dblSegmentStartX + (dblProportionAlongLine * (dblSegmentEndX - dblSegmentStartX))
  dblY_On_Segment = dblSegmentStartY + (dblProportionAlongLine * (dblSegmentEndY - dblSegmentStartY))
  
  Dim dblS As Double
  ' values interpreted as follows
  ' s<0      C is left of AB
  ' s>0      C is right of AB
  ' s=0      C is on AB
  dblS = (((dblSegmentStartY - dblPointY) * (dblSegmentEndX - dblSegmentStartX)) - _
         ((dblSegmentStartX - dblPointX) * (dblSegmentEndY - dblSegmentStartY))) / dblDenom
  
  If dblS < 0 Then
    lngClockwise = ENUM_CounterClockwise
  ElseIf dblS = 0 Then
    lngClockwise = Enum_OnLine
  Else
    lngClockwise = Enum_Clockwise
  End If
  
  dblDistToInfiniteLine = Abs(dblS) * (Sqr(dblDenom))
  
  If dblProportionAlongLine >= 0 And dblProportionAlongLine <= 1 Then
    DistancePointToSegment = dblDistToInfiniteLine
  Else
    Dim dblDistToStart As Double
    Dim dblDistToEnd As Double
    dblDistToStart = ((dblPointX - dblSegmentStartX) * (dblPointX - dblSegmentStartX)) + ((dblPointY - dblSegmentStartY) * (dblPointY - dblSegmentStartY))
    dblDistToEnd = ((dblPointX - dblSegmentEndX) * (dblPointX - dblSegmentEndX)) + ((dblPointY - dblSegmentEndY) * (dblPointY - dblSegmentEndY))
    If dblDistToStart < dblDistToEnd Then
      DistancePointToSegment = Sqr(dblDistToStart)
    Else
      DistancePointToSegment = Sqr(dblDistToEnd)
    End If
  End If

End Function


Public Function CalcFarthestPointsByNumbers(dblCoords() As Double, lngMethod As JenSphericalMethod, pPoint1 As IPoint, _
      pPoint2 As IPoint, pPointSpRef As ISpatialReference, dblDistance As Double, dblAZ1 As Double, dblAZ2 As Double, _
      dblReverseAz1 As Double, dblReverseAz2 As Double) As Boolean

  ' ACCEPTS GEOGRAPHIC OR PROJECTED DATA
  ' DON'T SEND THIS FUNCTION A NULL OR EMPTY GEOMETRY, OR ONE WITH ONLY ONE VERTEX.
  
  CalcFarthestPointsByNumbers = False
      
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngPointCount As Long
  Dim dblMaxDist As Double
  Dim dblTestDist As Double
    
  Dim dblTestAz1 As Double
  Dim dblTestAz2 As Double
  
  Dim dblStartX As Double
  Dim dblStartY As Double
  Dim dblEndX As Double
  Dim dblEndY As Double
  
  Dim dblFinalStartX As Double
  Dim dblFinalStartY As Double
  Dim dblFinalEndX As Double
  Dim dblFinalEndY As Double
  
  dblMaxDist = -999
  lngPointCount = UBound(dblCoords, 2)
'  Debug.Print CStr(lngPointCount) & " vertices..."
  If lngPointCount > 1 Then
    For lngIndex1 = 0 To lngPointCount - 2
      
      dblStartX = dblCoords(0, lngIndex1)
      dblStartY = dblCoords(1, lngIndex1)
      
      For lngIndex2 = lngIndex1 + 1 To lngPointCount - 1
      
        dblEndX = dblCoords(0, lngIndex2)
        dblEndY = dblCoords(1, lngIndex2)
        
        If lngMethod = ENUM_UseSpherical Then
          dblTestDist = DistanceHaversineNumbers(dblStartY, dblStartX, dblEndY, dblEndX, , True, dblTestAz1)
        ElseIf lngMethod = ENUM_UseSpheroidal Then
          dblTestDist = DistanceVincentyNumbers2(dblStartX, dblStartY, dblEndX, dblEndY, dblTestAz1, dblTestAz2)
        Else
          dblTestDist = (((dblStartX - dblEndX) ^ 2) + ((dblStartY - dblEndY) ^ 2)) ^ (0.5)
        End If
        
'        Debug.Print "Checking [" & CStr(Format(dblStartX, "0.000")) & ", " & CStr(Format(dblStartY, "0.000")) & "] to [" & _
              CStr(Format(dblEndX, "0.000")) & ", " & CStr(Format(dblEndY, "0.000")) & "]:  Distance = " & _
              CStr(Format(dblTestDist, "0")) & " meters..."
        
        If dblTestDist > dblMaxDist Then
          
'          Debug.Print "  --> Current Shortest Distance:  " & CStr(Format(dblTestDist, "0")) & " meters..."
          
          dblMaxDist = dblTestDist
          
          dblFinalStartX = dblStartX
          dblFinalStartY = dblStartY
          dblFinalEndX = dblEndX
          dblFinalEndY = dblEndY
          
          
'          Set pClone = pTestPoint1
'          Set pPoint1 = pClone.Clone
'          Set pClone = pTestPoint2
'          Set pPoint2 = pClone.Clone
          
          If lngMethod = ENUM_UseSpherical Then
            dblAZ1 = dblTestAz1
            If dblAZ1 > 360 Then dblAZ1 = dblAZ1 - 360
            If dblAZ1 < 0 Then dblAZ1 = dblAZ1 + 360
            dblAZ2 = dblAZ1
          ElseIf lngMethod = ENUM_UseSpheroidal Then
            dblAZ1 = dblTestAz1
            dblAZ2 = dblTestAz2
          Else
            dblAZ1 = CalcBearingNumbers(dblStartX, dblStartY, dblEndX, dblEndY)
            If dblAZ1 > 360 Then dblAZ1 = dblAZ1 - 360
            If dblAZ1 < 0 Then dblAZ1 = dblAZ1 + 360
            dblAZ2 = dblAZ1
          End If
          
'          Debug.Print "  --> Current Shortest Distance:  " & CStr(Format(dblTestDist, "0")) & " meters..."
'          Debug.Print "  --> [" & CStr(Format(pPoint1.X, "0.000")) & ", " & CStr(Format(pPoint1.Y, "0.000")) & "] to [" & _
'              CStr(Format(pPoint2.X, "0.000")) & ", " & CStr(Format(pPoint2.Y, "0.000")) & "]"
'          Debug.Print "  --> Current Azimuth:  " & CStr(Format(dblAz1, "0")) & " degrees..."
          
        End If
        
      Next lngIndex2
    Next lngIndex1
    
    dblDistance = dblMaxDist
    
    Set pPoint1 = New Point
    Set pPoint1.SpatialReference = pPointSpRef
    pPoint1.PutCoords dblFinalStartX, dblFinalStartY
    
    Set pPoint2 = New Point
    Set pPoint2.SpatialReference = pPointSpRef
    pPoint2.PutCoords dblFinalEndX, dblFinalEndY
    
    dblReverseAz1 = dblAZ1 - 180
    If dblReverseAz1 < 0 Then dblReverseAz1 = dblReverseAz1 + 360
    dblReverseAz2 = dblAZ2 - 180
    If dblReverseAz2 < 0 Then dblReverseAz2 = dblReverseAz2 + 360
    
    CalcFarthestPointsByNumbers = True
  End If

End Function

Public Function ProjectToWorldAzimuthal(pGeom As IGeometry, _
      Optional booFoundProblems As Boolean = False, Optional strProblemReason As String) As esriSystem.IArray
  
  ' RETURN ARRAY WILL CONTAIN 2 OBJECTS:
  ' 0) PROJECTED POLYGON
  ' 1) SPATIAL REFERENCE OBJECT
  
  Dim pOrigSpRef As ISpatialReference
  Set pOrigSpRef = pGeom.SpatialReference
  If Not TypeOf pOrigSpRef Is IGeographicCoordinateSystem Then
    booFoundProblems = True
    strProblemReason = "Spatial Reference Not Geographic"
    Exit Function
  End If
  
  Dim dblXMin As Double
  Dim dblXMax As Double
  Dim dblYMin As Double
  Dim dblYMax As Double
  
  Dim pPoint As IPoint
  Dim pEnv As IEnvelope
  
  Dim dblMeters As Double
  dblMeters = 100
  
  If TypeOf pGeom Is IPoint Then
    Set pPoint = pGeom
    dblXMin = pPoint.x - (dblMeters * 10)
    dblXMax = pPoint.x + (dblMeters * 10)
    dblYMin = pPoint.Y - (dblMeters * 10)
    dblYMax = pPoint.Y + (dblMeters * 10)
  Else
    Set pPoint = New Point
    Set pEnv = pGeom.Envelope
    pPoint.PutCoords (pEnv.XMax - pEnv.XMin) / 2 + pEnv.XMin, (pEnv.YMax - pEnv.YMin) / 2 + pEnv.YMin
    dblXMin = pEnv.XMin - (dblMeters * 10)
    dblXMax = pEnv.XMax + (dblMeters * 10)
    dblYMin = pEnv.YMin - (dblMeters * 10)
    dblYMax = pEnv.YMax + (dblMeters * 10)
  End If
    
  ' PROJECT INTO AZIMUTHAL EQUIDISTANT
  Dim pSpRefFact As ISpatialReferenceFactory3
  Set pSpRefFact = New SpatialReferenceEnvironment
  Dim pPrjCS As IProjectedCoordinateSystem3
  Set pPrjCS = pSpRefFact.CreateProjectedCoordinateSystem(esriSRProjCS_World_AzimuthalEquidistant)
  Dim pSpRef As ISpatialReference
  Set pSpRef = pPrjCS
  pPrjCS.CentralMeridian(True) = pPoint.x
  pPrjCS.LatitudeOfOrigin = pPoint.Y
  
  If Not MyGeomCheckSpRefDomain(pSpRef) Then
    Dim pSpRefRes As ISpatialReferenceResolution
    Set pSpRefRes = pSpRef
    pSpRefRes.ConstructFromHorizon
  End If
  
  pGeom.Project pSpRef
  Dim pTopoOp As ITopologicalOperator
  Set pTopoOp = pGeom
  pTopoOp.Simplify
  
  Set ProjectToWorldAzimuthal = New esriSystem.Array
  ProjectToWorldAzimuthal.Add pGeom
  ProjectToWorldAzimuthal.Add pSpRef
      Set pOrigSpRef = Nothing
      
  Set pPoint = Nothing
  Set pEnv = Nothing
  Set pSpRefFact = Nothing
  Set pPrjCS = Nothing
  Set pSpRef = Nothing
  Set pSpRefRes = Nothing
  Set pTopoOp = Nothing

End Function


Public Function ReturnLongestPerpendicularsFromSegment(dblCoordinates() As Double, dblStartX As Double, _
    dblStartY As Double, dblEndX As Double, dblEndY As Double, dblLengthClockwise As Double, _
    dblLengthCounterClockwise As Double, Optional dblFarCW_X As Double, Optional dblFarCW_Y As Double, _
    Optional dblFarCCW_X As Double, Optional dblFarCCW_Y As Double) As Boolean
    
  ' TREATS SEGMENT AS AN INFINITE LINE
  ' ASSUMES LINE IS PROJECTED!
    
  Dim lngIndex As Long
  Dim dblTestX As Double
  Dim dblTestY As Double
  Dim dblFarthestClockwise As Double
  Dim dblFarthestCCW As Double
  Dim dblTestDist As Double
  Dim lngClockwise As JenClockwiseConstants
  
  Dim pDebug As IPoint
  
  For lngIndex = 0 To UBound(dblCoordinates, 2)
    dblTestX = dblCoordinates(0, lngIndex)
    dblTestY = dblCoordinates(1, lngIndex)
    dblTestDist = DistancePointToInfiniteLine(dblStartX, dblStartY, dblEndX, dblEndY, dblTestX, dblTestY, lngClockwise)
    
 
    If lngClockwise = Enum_Clockwise Then
      If dblTestDist >= dblFarthestClockwise Then
        dblFarthestClockwise = dblTestDist
        dblFarCW_X = dblTestX
        dblFarCW_Y = dblTestY
      End If
    ElseIf lngClockwise = ENUM_CounterClockwise Then
      If dblTestDist >= dblFarthestCCW Then
        dblFarthestCCW = dblTestDist
        dblFarCCW_X = dblTestX
        dblFarCCW_Y = dblTestY
      End If
    ' SKIP POINTS THAT LAY ON LINE
    End If
    
  Next lngIndex
  
  dblLengthClockwise = dblFarthestClockwise
  dblLengthCounterClockwise = dblFarthestCCW
  
  ReturnLongestPerpendicularsFromSegment = True

End Function

Public Function ReturnWeightedMeanDir(dblCompassDirs() As Double) As Double

  Dim dblSumC As Double
  Dim dblSumS As Double
  Dim lngIndex As Long
  Dim dblRadians As Double
  Dim dblWeight As Double
  
'  Dim dblSumWeights As Double
  ' ASSUMES DIR IN 1ST COLUMN, WEIGHTS IN SECOND
  
  For lngIndex = 0 To UBound(dblCompassDirs, 2)
    
    dblRadians = AsRadians(dblCompassDirs(0, lngIndex))
    dblWeight = AsRadians(dblCompassDirs(1, lngIndex))
    dblSumC = dblSumC + (Cos(dblRadians) * dblWeight)
    dblSumS = dblSumS + (Sin(dblRadians) * dblWeight)
'    dblSumWeights = dblSumWeights + dblWeight
  Next lngIndex
  
'  Dim dblR As Double
'  dblR = Sqr(dblSumC ^ 2 + dblSumS ^ 2)
  
  Dim dblMeanDir As Double
  If Abs(dblSumC) < 0.00000001 And Abs(dblSumS) < 0.00000001 Then
    dblMeanDir = -9999
  Else
    dblMeanDir = atan2(dblSumS, dblSumC)
    dblMeanDir = AsDegrees(dblMeanDir)
    
    If dblMeanDir < 0 Then
      dblMeanDir = dblMeanDir + 360
    End If
  End If
  ReturnWeightedMeanDir = dblMeanDir

End Function


Public Function ReturnWeightedMeanDir2(dblCompassDirs() As Double, Optional dblMeanResultLength As Double, _
    Optional dblCircularVariance As Double, Optional dblAngularVariance As Double, _
    Optional dblCircularStandDev As Double, Optional dblAngularDeviation As Double, _
    Optional dblResultantLength As Double, Optional dblKappa As Double) As Double

  Dim dblSumC As Double
  Dim dblSumS As Double
  Dim lngIndex As Long
  Dim dblRadians As Double
  Dim dblWeight As Double
  Dim dblSumWeights As Double
  
  ' ASSUMES DIRECTION (DEGREES) IN 1ST COLUMN, WEIGHTS IN SECOND
  
  For lngIndex = 0 To UBound(dblCompassDirs, 2)
    
    dblRadians = AsRadians(dblCompassDirs(0, lngIndex))
    dblWeight = dblCompassDirs(1, lngIndex)
    dblSumC = dblSumC + (Cos(dblRadians) * dblWeight)
    dblSumS = dblSumS + (Sin(dblRadians) * dblWeight)
    dblSumWeights = dblSumWeights + dblWeight
  Next lngIndex
  
  
  Dim dblMeanDir As Double
  If Abs(dblSumC) < 0.00000001 And Abs(dblSumS) < 0.00000001 Then
    dblMeanDir = -9999
  Else
    dblMeanDir = atan2(dblSumS, dblSumC)
    dblMeanDir = AsDegrees(dblMeanDir)
    
    ForceAzimuthToCorrectRange dblMeanDir
    
    If dblMeanDir < 0 Then
      dblMeanDir = dblMeanDir + 360
    End If
  End If
  ReturnWeightedMeanDir2 = dblMeanDir

  dblResultantLength = Sqr(dblSumC ^ 2 + dblSumS ^ 2)
  dblMeanResultLength = dblResultantLength / dblSumWeights
  If dblMeanResultLength > 1 Then dblMeanResultLength = 1   ' ROUNDING ERROR CAN CAUSE THIS TO BE > 1 WHEN THERE IS NO VARIANCE
  dblCircularVariance = 1 - dblMeanResultLength
  dblAngularVariance = 2 * dblCircularVariance
  dblCircularStandDev = Sqr(-2 * (Log(dblMeanResultLength)))
  dblAngularDeviation = Sqr(dblAngularVariance)
  
  Dim lngPointCount As Long
  lngPointCount = UBound(dblCompassDirs, 2) + 1
  dblKappa = ReturnVonMisesKappa(dblMeanResultLength, lngPointCount, True)

End Function

Public Function ReturnVonMisesKappa(dblMeanResultLength As Double, lngPointCount As Long, booCorrectIfSmallSample As Boolean) As Double
  
  ' VON MISES DISPERSION:  KAPPA
  ' FROM FISHER, P. 88
  
  Dim dblKappa As Double
  If dblMeanResultLength < 0.53 Then
    dblKappa = (2 * dblMeanResultLength) + (dblMeanResultLength ^ 3) + (5 * (dblMeanResultLength ^ 5) / 6)
  ElseIf dblMeanResultLength < 0.85 Then
    dblKappa = -0.4 + (1.39 * dblMeanResultLength) + (0.43 / (1 - dblMeanResultLength))
  Else
    If ((dblMeanResultLength ^ 3) - (4 * (dblMeanResultLength ^ 2)) + (3 * dblMeanResultLength)) = 0 Then
      dblKappa = 1 / 0.000000001
    Else
      dblKappa = 1 / ((dblMeanResultLength ^ 3) - (4 * (dblMeanResultLength ^ 2)) + (3 * dblMeanResultLength))
    End If
  End If
  
  ' ADJUST KAPPA FOR SMALL SAMPLE SIZES
  If lngPointCount <= 15 And booCorrectIfSmallSample Then
    If dblKappa < 2 Then
      Dim dblTemp As Double
      dblTemp = dblKappa - (2 / (lngPointCount * dblKappa))
      If dblTemp < 0 Then
        dblKappa = 0
      Else
        dblKappa = dblTemp
      End If
    Else
      dblKappa = ((lngPointCount - 1) ^ 3) * dblKappa / (lngPointCount ^ 3 + lngPointCount)
    End If
  End If
  ReturnVonMisesKappa = dblKappa

End Function

Public Sub ForceAzimuthToCorrectRange(ByRef dblAz As Double)

  If dblAz < 0 Then
    Do Until dblAz > 0
      dblAz = dblAz + 360
    Loop
  End If
  
  If dblAz > 360 Then
    Do Until dblAz < 360
      dblAz = dblAz - 360
    Loop
  End If
  
  If dblAz = 360 Then dblAz = 0

End Sub
Public Sub ForceValueToCorrectRange(ByRef dblAz As Double, Optional dblMin As Double = 0, _
    Optional dblMax As Double = 360, Optional booMakeMaximumEqualMinimum As Boolean = True)
  
  Dim dblRange As Double
  dblRange = dblMax - dblMin
  
  If dblAz < dblMin Then
    Do Until dblAz > dblMin
      dblAz = dblAz + dblRange
    Loop
  End If
  
  If dblAz > dblMax Then
    Do Until dblAz < dblMax
      dblAz = dblAz - dblRange
    Loop
  End If
  
  If booMakeMaximumEqualMinimum Then
    If dblAz = dblMax Then dblAz = dblMin
  End If

End Sub



Public Function SquaredDistanceBetweenSegments( _
    dblSeg1Start() As Double, _
    dblSeg1End() As Double, _
    dblSeg2Start() As Double, _
    dblSeg2End() As Double, _
    dblClosePointOnSeg1() As Double, _
    dblClosePointOnSeg2() As Double) As Double
  
  Dim dblVectorU() As Double    ' VECTOR OF (SEGMENT 1 END POINT) - (SEGMENT 1 START POINT)
  Dim dblVectorV() As Double    ' VECTOR OF (SEGMENT 2 END POINT) - (SEGMENT 2 START POINT)
  Dim dblVectorW() As Double    ' VECTOR OF (SEGMENT 1 START POINT) - (SEGMENT 2 START POINT)
  
  Dim dblA As Double     ' DOT PRODUCT OF (VectorU * VectorU)
  Dim dblB As Double     ' DOT PRODUCT OF (VectorU * VectorV)
  Dim dblC As Double     ' DOT PRODUCT OF (VectorV * VectorV)
  Dim dblD As Double     ' DOT PRODUCT OF (VectorU * VectorW)
  Dim dblE As Double     ' DOT PRODUCT OF (VectorV * VectorW)
  Dim lngIndex As Long
  Dim lngUpper As Long
  Dim dblDenominator As Double
  Dim dblsc As Double
  Dim dblsN As Double
  Dim dblSD As Double
  Dim dblTC As Double
  Dim dbltN As Double
  Dim dbltD As Double
  
  Dim dblSmallNum As Double
  dblSmallNum = 0.000000000001
  
  lngUpper = UBound(dblSeg1Start)
  ReDim dblVectorU(lngUpper)
  ReDim dblVectorV(lngUpper)
  ReDim dblVectorW(lngUpper)
  ReDim dblClosePointOnSeg1(lngUpper)
  ReDim dblClosePointOnSeg2(lngUpper)
  
  dblA = 0
  dblB = 0
  dblC = 0
  dblD = 0
  dblE = 0
  
  For lngIndex = 0 To lngUpper
    dblVectorU(lngIndex) = (dblSeg1End(lngIndex) - dblSeg1Start(lngIndex))
    dblVectorV(lngIndex) = (dblSeg2End(lngIndex) - dblSeg2Start(lngIndex))
    dblVectorW(lngIndex) = (dblSeg1Start(lngIndex) - dblSeg2Start(lngIndex))
  Next lngIndex
  
  For lngIndex = 0 To lngUpper
    dblA = dblA + (dblVectorU(lngIndex) * dblVectorU(lngIndex))
    dblB = dblB + (dblVectorU(lngIndex) * dblVectorV(lngIndex))
    dblC = dblC + (dblVectorV(lngIndex) * dblVectorV(lngIndex))
    dblD = dblD + (dblVectorU(lngIndex) * dblVectorW(lngIndex))
    dblE = dblE + (dblVectorV(lngIndex) * dblVectorW(lngIndex))
  Next lngIndex
  
  dblDenominator = (dblA * dblC) - (dblB * dblB)
  dblsc = dblDenominator
  dblsN = dblDenominator
  dblSD = dblDenominator
  dblTC = dblDenominator
  dbltN = dblDenominator
  dbltD = dblDenominator
  
' Adapted from SoftSurfer code at http://softsurfer.com/Archive/algorithm_0106/algorithm_0106.htm#dist3D_Segment_to_Segment%28%29
'// dist3D_Segment_to_Segment():
'//    Input:  two 3D line segments S1 and S2
'//    Return: the shortest distance between S1 and S2
'Float
'dist3D_Segment_to_Segment( Segment S1, Segment S2)
'{
'    Vector   u = S1.P1 - S1.P0;
'    Vector   v = S2.P1 - S2.P0;
'    Vector   w = S1.P0 - S2.P0;
'    float    a = dot(u,u);        // always >= 0
'    float    b = dot(u,v);
'    float    c = dot(v,v);        // always >= 0
'    float    d = dot(u,w);
'    float    e = dot(v,w);
'    float    D = a*c - b*b;       // always >= 0
'    float    sc, sN, sD = D;      // sc = sN / sD, default sD = D >= 0
'    float    tc, tN, tD = D;      // tc = tN / tD, default tD = D >= 0
'

  If dblDenominator < dblSmallNum Then
    dblsN = 0
    dblSD = 1
    dbltN = dblE
    dbltD = dblC
  
  Else
    dblsN = (dblB * dblE) - (dblC * dblD)
    dbltN = (dblA * dblE) - (dblB * dblD)
    
    If dblsN < 0 Then
      dblsN = 0
      dbltN = dblE
      dbltD = dblC
    
    ElseIf dblsN > dblSD Then
      dblsN = dblSD
      dbltN = dblE + dblB
      dbltD = dblC
    End If
  End If
  
  If dbltN < 0 Then
    dbltN = 0
    
    If -dblD < 0 Then
      dblsN = 0
      
    ElseIf -dblD > dblA Then
      dblsN = dblSD
      
    Else
      dblsN = -dblD
      dblSD = dblA
      
    End If
  

    


'    // compute the line parameters of the two closest points
'    if (D < SMALL_NUM) { // the lines are almost parallel
'        sN = 0.0;        // force using point P0 on segment S1
'        sD = 1.0;        // to prevent possible division by 0.0 later
'        tN = e;
'        tD = c;
'    }
'    else {                // get the closest points on the infinite lines
'        sN = (b*e - c*d);
'        tN = (a*e - b*d);
'        if (sN < 0.0) {       // sc < 0 => the s=0 edge is visible
'            sN = 0.0;
'            tN = e;
'            tD = c;
'        }
'        else if (sN > sD) {  // sc > 1 => the s=1 edge is visible
'            sN = sD;
'            tN = e + b;
'            tD = c;
'        }
'    }
'
'    if (tN < 0.0) {           // tc < 0 => the t=0 edge is visible
'        tN = 0.0;
'        // recompute sc for this edge
'        if (-d < 0.0)
'            sN = 0.0;
'        else if (-d > a)
'            sN = sD;
'        else {
'            sN = -d;
'            sD = a;
'        }
'    }
    
  ElseIf dbltN > dbltD Then
    dbltN = dbltD
    
    If ((-dblD + dblB) < 0) Then
      dblsN = 0
      
    ElseIf ((-dblD + dblB) > dblA) Then
      dblsN = dblSD
      
    Else
      dblsN = -dblD + dblB
      dblSD = dblA
    
    End If
  End If
  
  If Abs(dblsN) < dblSmallNum Then
    dblsc = 0
  Else
    dblsc = dblsN / dblSD
  End If
  
  If Abs(dbltN) < dblSmallNum Then
    dblTC = 0
  Else
    dblTC = dbltN / dbltD
  End If
  
  
  
'    else if (tN > tD) {      // tc > 1 => the t=1 edge is visible
'        tN = tD;
'        // recompute sc for this edge
'        if ((-d + b) < 0.0)
'            sN = 0;
'        else if ((-d + b) > a)
'            sN = sD;
'        else {
'            sN = (-d + b);
'            sD = a;
'        }
'    }
'    // finally do the division to get sc and tc
'    sc = (abs(sN) < SMALL_NUM ? 0.0 : sN / sD);
'    tc = (abs(tN) < SMALL_NUM ? 0.0 : tN / tD);

'
'  For lngIndex = 0 To lngUpper
'    dblVectorU(lngIndex) = (dblSeg1End(lngIndex) - dblSeg1Start(lngIndex))
'    dblVectorV(lngIndex) = (dblSeg2End(lngIndex) - dblSeg2Start(lngIndex))
'    dblVectorW(lngIndex) = (dblSeg1Start(lngIndex) - dblSeg2Start(lngIndex))
'  Next lngIndex
  
  Dim dblP() As Double
  ReDim dblP(lngUpper)
  Dim dblDistance As Double
  dblDistance = 0
  For lngIndex = 0 To lngUpper
'    dblP(lngIndex) = (dblVectorW(lngIndex) + (dblsc * (dblVectorU(lngIndex))) - _
          (dbltc * (dblVectorV(lngIndex))))
    dblClosePointOnSeg1(lngIndex) = dblSeg1Start(lngIndex) + dblsc * (dblVectorU(lngIndex))
    dblClosePointOnSeg2(lngIndex) = dblSeg2Start(lngIndex) + dblTC * (dblVectorV(lngIndex))
    dblDistance = dblDistance + ((dblClosePointOnSeg1(lngIndex) - dblClosePointOnSeg2(lngIndex)) ^ 2)
  Next lngIndex
  
  SquaredDistanceBetweenSegments = dblDistance
'
'    // get the difference of the two closest points
'    Vector   dP = w + (sc * u) - (tc * v);  // = S1(sc) - S2(tc)
'
'    return norm(dP);   // return the closest distance
'}

End Function


Public Function SpheroidalPolylineFromEndPoints(pStartPoint As IPoint, pEndPoint As IPoint, _
    lngNumberVertices As Long)
  
  
  ' ASSUMES POINTS ARE IN GEOGRAPHIC COORDINATES!
  ' WILL USE DATUM OF POINTS TO GET EQUATORIAL AND POLAR RADIUS.
  
  Dim pSpRef As ISpatialReference
  Set pSpRef = pStartPoint.SpatialReference
  Dim pGeoSpRef As IGeographicCoordinateSystem
  If Not TypeOf pSpRef Is IGeographicCoordinateSystem Then
    Set SpheroidalPolylineFromEndPoints = Nothing
    GoTo ClearMemory
  End If
  
  Set pGeoSpRef = pSpRef
  Dim dblEquatorialRadius As Double
  Dim dblPolarRadius As Double
  dblEquatorialRadius = pGeoSpRef.Datum.Spheroid.SemiMajorAxis
  dblPolarRadius = pGeoSpRef.Datum.Spheroid.SemiMinorAxis
  
  Dim pInitialPolyline As IPointCollection
  Dim pGeom As IGeometry
  
  Set pInitialPolyline = New Polyline
  Set pGeom = pInitialPolyline
  Set pGeom.SpatialReference = pSpRef
  
  If pStartPoint.x = pEndPoint.x And pStartPoint.Y = pEndPoint.Y Then
    Set SpheroidalPolylineFromEndPoints = pInitialPolyline
    GoTo ClearMemory
  End If
  
  pInitialPolyline.AddPoint pStartPoint
  pInitialPolyline.AddPoint pEndPoint
  
  Dim pFinalPolyline As IPointCollection
  Set pFinalPolyline = New Polyline
  Set pGeom = pFinalPolyline
  Set pGeom.SpatialReference = pSpRef
  pFinalPolyline.AddPoint pStartPoint
  
  Dim dblIndex As Double
  Dim dblInterval As Double
  dblInterval = 1 / (lngNumberVertices - 1)
  
  Dim pPoint As IPoint
  
  For dblIndex = dblInterval To (1 - dblInterval) Step dblInterval
    Set pPoint = SpheroidalPolylineMidpoint2(pInitialPolyline, dblIndex, True, , dblEquatorialRadius, _
        dblPolarRadius)
    pFinalPolyline.AddPoint pPoint
  Next dblIndex
  
  pFinalPolyline.AddPoint pEndPoint
  Set SpheroidalPolylineFromEndPoints = pFinalPolyline
  
  Exit Function
ClearMemory:
  Set pSpRef = Nothing
  Set pGeoSpRef = Nothing
  Set pInitialPolyline = Nothing
  Set pGeom = Nothing
  Set pFinalPolyline = Nothing
  Set pPoint = Nothing

End Function

Public Function DegToPercent(dblDeg As Double) As Double
  
  DegToPercent = Tan(dblDeg * dblPI / 180)

End Function
Public Function PercentToDeg(dblPercent As Double) As Double
  
  PercentToDeg = Atn(dblPercent) * 180 / dblPI

End Function

Public Function CalcProjectedDistance(pPoint1 As IPoint, pPoint2 As IPoint) As Double
  
  If pPoint1.IsEmpty Then
    CalcProjectedDistance = -9999
  ElseIf pPoint2.IsEmpty Then
    CalcProjectedDistance = -9999
  Else
    CalcProjectedDistance = ((pPoint1.x - pPoint2.x) ^ 2 + (pPoint1.Y - pPoint2.Y) ^ 2) ^ (0.5)
  End If

End Function
Public Function CalcProjectedDistanceNumbers(dblX1 As Double, dblY1 As Double, dblX2 As Double, dblY2 As Double) As Double

  CalcProjectedDistanceNumbers = ((dblX1 - dblX2) ^ 2 + (dblY1 - dblY2) ^ 2) ^ (0.5)

End Function
Public Function UnionGeometries2(pGeomArray As esriSystem.IVariantArray, _
    Optional lngMaxNumberToUnion As Long = -999) As IGeometry
  
'  Dim pMxDox As IMxDocument
'  Set pMxDoc = ThisDocument
  
  Dim pTopoOp As ITopologicalOperator
  Dim pGeom As IGeometry
  Dim pGeometryCollection As IGeometryCollection
  
  Set pGeometryCollection = New GeometryBag
  
  Set pGeom = pGeomArray.Element(0)
  Dim pSpRef As ISpatialReference
  Dim pTempGeom As IGeometry
  Dim pNewGeom As IGeometry
  Set pSpRef = pGeom.SpatialReference
  
  Dim lngGeomType As esriGeometryType
  lngGeomType = pGeom.GeometryType
  
  Dim lngIndex As Long
  For lngIndex = 0 To pGeomArray.Count - 1
    Set pGeom = pGeomArray.Element(lngIndex)
    
    
    If Not pGeom.IsEmpty Then
      pGeometryCollection.AddGeometry pGeom
    End If
    
    If lngMaxNumberToUnion > 1 Then
      If pGeometryCollection.GeometryCount >= lngMaxNumberToUnion Then

        If lngGeomType = esriGeometryPoint Then
          Set pTempGeom = New Multipoint
        ElseIf lngGeomType = esriGeometryMultipoint Then
          Set pTempGeom = New Multipoint
        ElseIf lngGeomType = esriGeometryPolyline Then
          Set pTempGeom = New Polyline
        ElseIf lngGeomType = esriGeometryPolygon Then
          Set pTempGeom = New Polygon
        End If
        
        Set pTopoOp = pTempGeom
        pTopoOp.ConstructUnion pGeometryCollection
        pTopoOp.Simplify
        
        Set pTempGeom.SpatialReference = pSpRef
        Set pGeometryCollection = New GeometryBag
        pGeometryCollection.AddGeometry pTempGeom
        
      End If
    End If
    
  Next lngIndex
  
  If pGeometryCollection.GeometryCount = 1 Then
    Set pNewGeom = pGeometryCollection.Geometry(0)
  Else
    If lngGeomType = esriGeometryPoint Then
      Set pNewGeom = New Multipoint
    ElseIf lngGeomType = esriGeometryMultipoint Then
      Set pNewGeom = New Multipoint
    ElseIf lngGeomType = esriGeometryPolyline Then
      Set pNewGeom = New Polyline
    ElseIf lngGeomType = esriGeometryPolygon Then
      Set pNewGeom = New Polygon
    End If
    
    Set pTopoOp = pNewGeom
    pTopoOp.ConstructUnion pGeometryCollection
    pTopoOp.Simplify
    
    Set pNewGeom.SpatialReference = pSpRef
  End If
  
  Set UnionGeometries2 = pNewGeom
    
  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing
  
  GoTo ClearMemory
ClearMemory:
    
  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing

End Function
Public Function UnionGeometries3(pGeomArray As esriSystem.IVariantArray, _
    Optional lngMaxNumberToUnion As Long = -999) As IGeometry

'  Dim pMxDox As IMxDocument
'  Set pMxDoc = ThisDocument
  
  Dim pTopoOp As ITopologicalOperator
  Dim pGeom As IGeometry
  Dim pGeometryCollection As IGeometryCollection
  
  Set pGeometryCollection = New GeometryBag
  
  Dim pSpRef As ISpatialReference
  Dim pTempGeom As IGeometry
  Dim pNewGeom As IGeometry
  Dim lngIndex As Long
  Dim booFoundGeometry As Boolean
  
  Do Until lngIndex = pGeomArray.Count Or Not pSpRef Is Nothing
    Set pGeom = pGeomArray.Element(0)
    If Not pGeom Is Nothing Then
      Set pSpRef = pGeom.SpatialReference
      booFoundGeometry = True
    End If
    lngIndex = lngIndex + 1
  Loop
  
  Dim lngGeomType As esriGeometryType
  lngGeomType = pGeom.GeometryType
  
  If Not booFoundGeometry Then
    Set UnionGeometries3 = Nothing
  Else
    For lngIndex = 0 To pGeomArray.Count - 1
      Set pGeom = pGeomArray.Element(lngIndex)
      
      If Not pGeom Is Nothing Then
        If Not pGeom.IsEmpty Then
          pGeometryCollection.AddGeometry pGeom
      
          If lngMaxNumberToUnion > 1 Then
            If pGeometryCollection.GeometryCount >= lngMaxNumberToUnion Then
      
              If lngGeomType = esriGeometryPoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryMultipoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryPolyline Then
                Set pTempGeom = New Polyline
              ElseIf lngGeomType = esriGeometryPolygon Then
                Set pTempGeom = New Polygon
              End If
              
              Set pTopoOp = pTempGeom
              pTopoOp.ConstructUnion pGeometryCollection
              pTopoOp.Simplify
              
              Set pTempGeom.SpatialReference = pSpRef
              Set pGeometryCollection = New GeometryBag
              pGeometryCollection.AddGeometry pTempGeom
              
            End If
          End If
        End If
      End If
      
    Next lngIndex
    
    If pGeometryCollection.GeometryCount = 1 Then
      Set pNewGeom = pGeometryCollection.Geometry(0)
    Else
      If lngGeomType = esriGeometryPoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryMultipoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryPolyline Then
        Set pNewGeom = New Polyline
      ElseIf lngGeomType = esriGeometryPolygon Then
        Set pNewGeom = New Polygon
      End If
      
      Set pTopoOp = pNewGeom
      pTopoOp.ConstructUnion pGeometryCollection
      pTopoOp.Simplify
      
      Set pNewGeom.SpatialReference = pSpRef
    End If
    
    Set UnionGeometries3 = pNewGeom
  End If
  
  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing
  
  
  GoTo ClearMemory
ClearMemory:
    
  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing

End Function
Public Function UnionGeometries4(pGeomArray As esriSystem.IVariantArray, _
    Optional lngMaxNumberToUnion As Long = -999) As IGeometry

'  Dim pMxDox As IMxDocument
'  Set pMxDoc = ThisDocument
  
  Dim pTopoOp As ITopologicalOperator
  Dim pGeom As IGeometry
  Dim pGeometryCollection As IGeometryCollection
  
  Set pGeometryCollection = New GeometryBag
  
  Dim pSpRef As ISpatialReference
  Dim pTempGeom As IGeometry
  Dim pNewGeom As IGeometry
  Dim lngIndex As Long
  Dim booFoundGeometry As Boolean
  
  Do Until lngIndex = pGeomArray.Count Or Not pSpRef Is Nothing
    Set pGeom = pGeomArray.Element(0)
    If Not pGeom Is Nothing Then
      Set pSpRef = pGeom.SpatialReference
      booFoundGeometry = True
    End If
    lngIndex = lngIndex + 1
  Loop
  
  Dim lngGeomType As esriGeometryType
  lngGeomType = pGeom.GeometryType
  
  Dim pTempPoly As IPolygon
  Dim pNewPoly As IPolygon
  Dim pSimplifyTopoOp As ITopologicalOperator4
  
'  Dim pClone As IClone
'  Dim pTempEnv As IEnvelope
'  Dim pTempPoly2 As IPolygon
'  Set pClone = pGeomArray.Element(1)
'  Set pTempPoly = pClone.Clone
'  Set pTempEnv = pTempPoly.Envelope
'  Set pTempPoly = MyGeometricOperations.EnvelopeToPolygon(pTempEnv)
'
'  Set pTopoOp = pTempPoly
'  Set pNewPoly = pClone.Clone
'  Set pTempPoly2 = pTopoOp.Union(pNewPoly)
  
  If Not booFoundGeometry Then
    Set UnionGeometries4 = Nothing
  Else
    For lngIndex = 0 To pGeomArray.Count - 1
      Set pGeom = pGeomArray.Element(lngIndex)
      
      If Not pGeom Is Nothing Then
        If Not pGeom.IsEmpty Then
          Set pSimplifyTopoOp = pGeom
          pSimplifyTopoOp.IsKnownSimple = False
          pSimplifyTopoOp.Simplify
          
          pGeometryCollection.AddGeometry pGeom
      
          If lngMaxNumberToUnion > 1 Then
            If pGeometryCollection.GeometryCount >= lngMaxNumberToUnion Then
      
              If lngGeomType = esriGeometryPoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryMultipoint Then
                Set pTempGeom = New Multipoint
              ElseIf lngGeomType = esriGeometryPolyline Then
                Set pTempGeom = New Polyline
              ElseIf lngGeomType = esriGeometryPolygon Then
                Set pTempGeom = New Polygon
              End If
              
              Set pTopoOp = pTempGeom
              pTopoOp.ConstructUnion pGeometryCollection
              pTopoOp.Simplify
              
              Set pTempGeom.SpatialReference = pSpRef
              Set pGeometryCollection = New GeometryBag
              pGeometryCollection.AddGeometry pTempGeom
              
            End If
          End If
        End If
      End If
      
    Next lngIndex
      
    If pGeometryCollection.GeometryCount = 1 Then
      Set pNewGeom = pGeometryCollection.Geometry(0)
      
'    ElseIf pGeometryCollection.GeometryCount <= 4 Then
'      Set pNewPoly = pGeometryCollection.Geometry(0)
'      Set pTopoOp = pNewPoly
'      For lngIndex = 1 To pGeometryCollection.GeometryCount - 1
'        Set pTempPoly = pGeometryCollection.Geometry(lngIndex)
'        Set pNewPoly = pTopoOp.Union(pTempPoly)
'        Set pTopoOp = pNewPoly
'      Next lngIndex
      
    Else
      If lngGeomType = esriGeometryPoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryMultipoint Then
        Set pNewGeom = New Multipoint
      ElseIf lngGeomType = esriGeometryPolyline Then
        Set pNewGeom = New Polyline
      ElseIf lngGeomType = esriGeometryPolygon Then
        Set pNewGeom = New Polygon
      End If
      
      Set pTopoOp = pNewGeom
      pTopoOp.ConstructUnion pGeometryCollection
      pTopoOp.Simplify
      
      Set pNewGeom.SpatialReference = pSpRef
    End If
    
    Set UnionGeometries4 = pNewGeom
  End If
  
  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing
  
  
  GoTo ClearMemory
ClearMemory:
    
  Set pTopoOp = Nothing
  Set pGeom = Nothing
  Set pGeometryCollection = Nothing
  Set pSpRef = Nothing
  Set pNewGeom = Nothing
  Set pTempGeom = Nothing

End Function

Public Function CalcDirectionDeviationDegrees(dblAngle1 As Double, dblAngle2 As Double) As Double
  
  ' GIVES THE DIFFERENCE IN DEGREES BETWEEN ANGLE 1 AND ANGLE 2.  POSITIVE IF ANGLE 2 IS CLOCKWISE
  ' FROM ANGLE 1.
  
  CalcDirectionDeviationDegrees = MyGeometricOperations.ModDouble(Abs(dblAngle2 - dblAngle1), 360)
  If CalcDirectionDeviationDegrees > 180 Then CalcDirectionDeviationDegrees = 360 - CalcDirectionDeviationDegrees
  
  Dim dblPX As Double
  Dim dblPY As Double
  Dim dblQX As Double
  Dim dblQY As Double
  Dim dblRX As Double
  Dim dblRY As Double
  
  dblPX = 0
  dblPY = 0
  
  MyGeometricOperations.CalcPointNumbers dblPX, dblPY, 1, dblAngle1, dblQX, dblQY
  MyGeometricOperations.CalcPointNumbers dblQX, dblQY, 1, dblAngle2, dblRX, dblRY
  
  Dim booClockwise As Boolean
  booClockwise = MyGeometricOperations.CalcCheckClockwiseNumbers(dblPX, dblPY, dblQX, dblQY, dblRX, dblRY)
  
  If Not booClockwise Then CalcDirectionDeviationDegrees = -(Abs(CalcDirectionDeviationDegrees))

End Function
Public Sub CalcPointNumbers(dblOriginX As Double, dblOriginY As Double, theLength As Double, _
  dblAzimuth As Double, dblEndPointX As Double, dblEndPointY As Double)
  
  ' Jenness Enterprises <www.jennessent.com>
  ' Given an origin point, distance and bearing, this script will return a new point at that distance and bearing, and a line
  ' connecting that new point to the origin point
  
  '' MAKE SURE AZIMUTH IS BETWEEN 0 AND 360
  Dim theAzimuth As Double
  theAzimuth = dblAzimuth
  
  theAzimuth = ModDouble(theAzimuth, 360)
  
  'theAzimuth = theAzimuth Mod 360
  '
  '' NEW SEGMENT AND POINT DISTANCE NORTH/SOUTH AND EAST/WEST BASED ON DISTANCE AND BEARING FROM ORIGIN.
  '' THERE ARE EIGHT DIFFERENT POSSIBILITIES:  THE BEARING COULD BE ONE OF THE FOUR CARDINAL DIRECTIONS OR IT
  '' COULD BE IN ONE OF THE FOUR QUADRANTS.  THE BEARING IS TREATED DIFFERENTLY IN EACH OF THESE CIRCUMSTANCES.
  '' THE CALCULATION TO DETERMINE THE NEW POINT LOCATION IS ESSENTIALLY A MATTER OF TAKING THE SINE OR THE
  '' COSINE OF THE ANGLE (AFTER CONVERTING IT TO RADIANS), AND MULTIPLYING THAT SINE OR COSINE BY THE MEASURED
  '' DISTANCE.  PLEASE CONTACT THE AUTHOR IF THIS DOESN'T MAKE SENSE, OR IF YOU WOULD LIKE FURTHER EXPLANATION.
  Dim NorthSouthDistance As Double
  Dim EastWestDistance As Double
  Dim EastWest As Integer
  Dim NorthSouth As Integer
  
  If theAzimuth = 0 Or theAzimuth = 360 Then
    NorthSouthDistance = theLength
    NorthSouth = 1
    EastWestDistance = 0
    EastWest = 1
  ElseIf (theAzimuth = 180) Then
    NorthSouthDistance = theLength
    NorthSouth = -1
    EastWestDistance = 0
    EastWest = 1
  ElseIf (theAzimuth = 90) Then
    NorthSouthDistance = 0
    NorthSouth = 1
    EastWestDistance = theLength
    EastWest = 1
  ElseIf (theAzimuth = 270) Then
    NorthSouthDistance = 0
    NorthSouth = 1
    EastWestDistance = theLength
    EastWest = -1
  ElseIf ((theAzimuth > 0) And (theAzimuth < 90)) Then
    NorthSouthDistance = Cos(AsRadians(theAzimuth)) * theLength
    NorthSouth = 1
    EastWestDistance = Sin(AsRadians(theAzimuth)) * theLength
    EastWest = 1
  ElseIf ((theAzimuth > 90) And (theAzimuth < 180)) Then
    NorthSouthDistance = (Sin(AsRadians(theAzimuth - 90))) * theLength
    NorthSouth = -1
    EastWestDistance = (Cos(AsRadians(theAzimuth - 90))) * theLength
    EastWest = 1
  ElseIf ((theAzimuth > 180) And (theAzimuth < 270)) Then
    NorthSouthDistance = (Cos(AsRadians(theAzimuth - 180))) * theLength
    NorthSouth = -1
    EastWestDistance = (Sin(AsRadians(theAzimuth - 180))) * theLength
    EastWest = -1
  ElseIf ((theAzimuth > 270) And (theAzimuth < 360)) Then
    NorthSouthDistance = (Sin(AsRadians(theAzimuth - 270))) * theLength
    NorthSouth = 1
    EastWestDistance = (Cos(AsRadians(theAzimuth - 270))) * theLength
    EastWest = -1
  End If
  
  Dim theMovementNorth As Double
  Dim theMovementWest As Double
  
  theMovementNorth = NorthSouthDistance * NorthSouth
  theMovementWest = EastWestDistance * EastWest
  
  dblEndPointX = dblOriginX + theMovementWest
  dblEndPointY = dblOriginY + theMovementNorth

End Sub


Public Function TriangulatePolygonToDouble(pPolygon As IPolygon) As Double()
  
  ' double array should have 7 columns, 1 for relative proportional area of triangle and 6 for three point coordinates
  Dim pPolyHelper As ILinePolygonHelper
  Set pPolyHelper = New LinePolygonHelper
  
  Dim pMultiPatch As IMultiPatch
  Set pMultiPatch = New MultiPatch
  Dim booSuccess As Boolean
  booSuccess = pPolyHelper.Triangulate(pPolygon, pMultiPatch)
  
'  Dim pTriangles As ITriangles
'  Set pTriangles = pMultiPatch
  
  Dim pGeomColl As IGeometryCollection
  Set pGeomColl = pMultiPatch
  
  Dim pPtColl As IPointCollection
  Set pPtColl = pMultiPatch
  
  Dim dblX As Double
  Dim dblY As Double
  
  Dim pPoint As IPoint
  Set pPoint = New Point
  
  Dim dblCoords() As Double
  ReDim dblCoords(6, (pPtColl.PointCount / 3) - 1)
  
  Dim dblX1 As Double
  Dim dblY1 As Double
  Dim dblX2 As Double
  Dim dblY2 As Double
  Dim dblX3 As Double
  Dim dblY3 As Double
  
  Dim lngIndex As Long
  Dim lngTriangleIndex As Long
  
  Dim dblTriangleArea As Double
  Dim dblCumulativeArea As Double
  
  dblCumulativeArea = 0
  
  lngTriangleIndex = -1
  For lngIndex = 0 To pPtColl.PointCount - 1 Step 3
    pPtColl.QueryPoint lngIndex, pPoint
    dblX1 = pPoint.x
    dblY1 = pPoint.Y
    pPtColl.QueryPoint lngIndex + 1, pPoint
    dblX2 = pPoint.x
    dblY2 = pPoint.Y
    pPtColl.QueryPoint lngIndex + 2, pPoint
    dblX3 = pPoint.x
    dblY3 = pPoint.Y
    
    dblTriangleArea = TriangleAreaPointsValues(dblX1, dblY1, dblX2, dblY2, dblX3, dblY3)
    dblCumulativeArea = dblCumulativeArea + dblTriangleArea
    
    lngTriangleIndex = lngTriangleIndex + 1
    dblCoords(0, lngTriangleIndex) = dblCumulativeArea
    dblCoords(1, lngTriangleIndex) = dblX1
    dblCoords(2, lngTriangleIndex) = dblY1
    dblCoords(3, lngTriangleIndex) = dblX2
    dblCoords(4, lngTriangleIndex) = dblY2
    dblCoords(5, lngTriangleIndex) = dblX3
    dblCoords(6, lngTriangleIndex) = dblY3
    
  Next lngIndex
  
  For lngIndex = 0 To UBound(dblCoords, 2)
    
    dblTriangleArea = dblCoords(0, lngIndex)
    dblCoords(0, lngIndex) = dblTriangleArea / dblCumulativeArea
  Next lngIndex
  
  TriangulatePolygonToDouble = dblCoords
  
  GoTo ClearMemory
  
ClearMemory:
  Set pPolyHelper = Nothing
  Set pMultiPatch = Nothing
  Set pGeomColl = Nothing
  Set pPtColl = Nothing
  Set pPoint = Nothing
  Erase dblCoords

End Function
Public Function RandomlySelectTriangle(dblCoordArray() As Double) As Long

  ' RETURNS THE INDEX VALUE FOR THE ARRAY.  ASSUMES THE FIRST COLUMN IS A CUMULATIVE PROPORTION.
  ' ADAPTED FROM http://www.freevbcode.com/ShowCode.asp?ID=9416
  
  Dim dblRandom As Double
  dblRandom = Rnd()
    
  Dim low As Long
  low = 0
  Dim high As Long
  high = UBound(dblCoordArray, 2)
  Dim i As Long
  Dim result As Boolean
  Dim booFound As Boolean
  Dim dblLowRange As Double
  Dim dblHighRange As Double
  
  Do While low <= high
    i = (low + high) / 2
    If i = 0 Then
      dblLowRange = 0
    Else
      dblLowRange = dblCoordArray(0, i - 1)
    End If
    dblHighRange = dblCoordArray(0, i)
    
    booFound = dblRandom >= dblLowRange And dblRandom <= dblHighRange
    
    If booFound Then
        RandomlySelectTriangle = i
        Exit Do
    ElseIf dblRandom < dblLowRange Then
        high = (i - 1)
    Else
        low = (i + 1)
    End If
  Loop
    
'  Dim lngIndex As Long
'  For lngIndex = 0 To UBound(dblCoordArray, 2)
'    If dblCoordArray(0, lngIndex) > dblRandom Then
'      RandomlySelectTriangle = lngIndex
'      Exit For
'    End If
'  Next lngIndex
'
'  Debug.Print "i = " & CStr(i)
'  Debug.Print "lngIndex = " & CStr(lngIndex)
'  Debug.Print

End Function

Public Function RandomPointInTriangle(dblTriX1 As Double, dblTriY1 As Double, _
    dblTriX2 As Double, dblTriY2 As Double, dblTriX3 As Double, dblTriY3 As Double, _
    dblRandomX As Double, dblRandomY As Double) As Boolean
    
  RandomPointInTriangle = False
  
  Dim dblRandom1 As Double
  Dim dblRandom2 As Double
  dblRandom1 = Rnd()
  dblRandom2 = Rnd()
  
  Do Until dblRandom1 + dblRandom2 < 1
    dblRandom1 = Rnd()
    dblRandom2 = Rnd()
  Loop
  
  dblRandomX = ((dblTriX2 - dblTriX1) * dblRandom1) + ((dblTriX3 - dblTriX1) * dblRandom2) + dblTriX1
  dblRandomY = ((dblTriY2 - dblTriY1) * dblRandom1) + ((dblTriY3 - dblTriY1) * dblRandom2) + dblTriY1
  
  RandomPointInTriangle = True

End Function

Public Function CheckPointInTriangle(dblTriX1 As Double, dblTriY1 As Double, _
    dblTriX2 As Double, dblTriY2 As Double, dblTriX3 As Double, dblTriY3 As Double, _
    dblTestPointX As Double, dblTestPointY As Double) As Boolean

' ADAPTED FROM http://stackoverflow.com/questions/2049582/how-to-determine-a-point-in-a-triangle
  
  Dim boo1 As Boolean
  Dim boo2 As Boolean
  Dim boo3 As Boolean
  
  boo1 = PointInTriangleSign(dblTestPointX, dblTestPointY, dblTriX1, dblTriY1, dblTriX2, dblTriY2)
  boo2 = PointInTriangleSign(dblTestPointX, dblTestPointY, dblTriX2, dblTriY2, dblTriX3, dblTriY3)
  boo3 = PointInTriangleSign(dblTestPointX, dblTestPointY, dblTriX3, dblTriY3, dblTriX1, dblTriY1)
  
  CheckPointInTriangle = (boo1 = boo2) And (boo2 = boo3)

End Function

Private Function PointInTriangleSign(dblTestPointX As Double, dblTestPointY As Double, _
    dblTriX1 As Double, dblTriY1 As Double, dblTriX2 As Double, dblTriY2 As Double) As Double

  ' ADAPTED FROM http://stackoverflow.com/questions/2049582/how-to-determine-a-point-in-a-triangle
  Dim dblTest As Double
  dblTest = ((dblTestPointX - dblTriX2) * (dblTriY1 - dblTriY2)) - _
                        ((dblTriX1 - dblTriX2) * (dblTestPointY - dblTriY2))
  PointInTriangleSign = dblTest < 0

End Function

Public Function RandomPointInPolygon(dblPolyArray() As Double, dblRandX As Double, dblRandY As Double) As Boolean
    
  Dim lngThresholdCounter As Long
  Dim lngThresholdCount As Long
  lngThresholdCount = 1000
  
  RandomPointInPolygon = False
  Dim lngTriangleIndex As Long
  Dim booTest As Boolean
  
  lngTriangleIndex = MyGeometricOperations.RandomlySelectTriangle(dblPolyArray)
  booTest = False
  Do Until booTest Or lngThresholdCounter > lngThresholdCount
    booTest = MyGeometricOperations.RandomPointInTriangle(dblPolyArray(1, lngTriangleIndex), _
        dblPolyArray(2, lngTriangleIndex), dblPolyArray(3, lngTriangleIndex), dblPolyArray(4, lngTriangleIndex), _
        dblPolyArray(5, lngTriangleIndex), dblPolyArray(6, lngTriangleIndex), dblRandX, dblRandY)
    lngThresholdCounter = lngThresholdCounter + 1
  Loop
  
  RandomPointInPolygon = True

ClearMemory:

End Function

Public Function ConvertSlopeDegreesToPercent(dblDegrees As Double) As Double

  ConvertSlopeDegreesToPercent = Tan(AsRadians(dblDegrees)) * 100

End Function

Public Function ConvertSlopePercentToDegrees(dblPercentSlope As Double) As Double

  ConvertSlopePercentToDegrees = AsDegrees(Atn(dblPercentSlope / 100))

End Function

Public Function ReturnPi() As Double

  ' FROM http://mathworld.wolfram.com/PiFormulas.html
  ' BASED ON MACHIN'S FORMULA
  ReturnPi = (4 * Atn(1 / 5) - Atn(1 / 239)) * 4

End Function

Public Function SplitGeometryOnDateLine(pPolygonOrPolyline As IGeometry, booSucceeded As Boolean, _
  strReasonForFailure As String) As IGeometry

  Dim pGeoPolyline As IPolyline
  Dim pGeoPolygon As IPolygon
  Dim pPolyline As IPolyline
  Dim pPolygon As IPolygon
  Dim pClone As IClone
  Dim pPrjSpRef As IProjectedCoordinateSystem
  Dim pGeoSpRef As IGeographicCoordinateSystem
  Dim pSpRef As ISpatialReference
  Dim pPoint As IPoint
  Dim pNewPoint As IPoint
  Dim pNewPolyline As IPolyline
  Dim dblAZ2 As Double
      
  Dim pGeomColl As IGeometryCollection
  Dim lngIndex As Long
  Dim pSubPolygon As IPolygon
  Dim pSubEnv As IEnvelope
  Dim pNewPolygon As IPolygon
    
  Dim pSplitPolyline As IPolyline
  Dim pSplitPtColl As IPointCollection
  Dim pSplitPoint As IPoint
  Dim lngSplitIndex As Long
  Dim pTopoOp As ITopologicalOperator2
  Dim pTopoCutter As ITopologicalOperator4
  Dim pLeft As IPolyline
  Dim pRight As IPolyline
  
  Dim pTransform As ITransform2D
  Dim pCombineArray As IVariantArray
  
  Dim pEnv As IEnvelope
  
  booSucceeded = True
  
  Set pSpRef = pPolygonOrPolyline.SpatialReference
  If pSpRef Is Nothing Then
    booSucceeded = False
    strReasonForFailure = "No coordinate syatem available!"
  ElseIf TypeOf pSpRef Is IUnknownCoordinateSystem Then
    booSucceeded = False
    strReasonForFailure = "Unknown coordinate syatem!"
  End If
  
  
  If booSucceeded Then
    
    If TypeOf pPolygonOrPolyline Is IPolyline Then
      Set pPolyline = pPolygonOrPolyline
      Set pClone = pPolyline
      Set pGeoPolyline = pClone.Clone
      
      Set pSpRef = pPolyline.SpatialReference
      
      If Not TypeOf pSpRef Is IGeographicCoordinateSystem Then
        Set pPrjSpRef = pSpRef
        Set pGeoSpRef = pPrjSpRef.GeographicCoordinateSystem
        pGeoPolyline.Project pGeoSpRef
      End If
      
      ' CHECK TO SEE IF THIS NEEDS TO BE SPLIT AT ALL
      Set pEnv = pGeoPolyline.Envelope
      If pEnv.XMin < -180 Or pEnv.XMax > 180 Then
   
  '      Debug.Print "Min X = " & pNewPolyline.Envelope.XMin
  '      Debug.Print "Max X = " & pNewPolyline.Envelope.XMax
              
        If pEnv.XMin < -180 Then
        
          Set pSplitPolyline = New Polyline
          Set pSplitPolyline.SpatialReference = pGeoPolyline.SpatialReference
          Set pSplitPtColl = pSplitPolyline
          For lngSplitIndex = -80 To 80 Step 5
            Set pSplitPoint = New Point
            Set pSplitPoint.SpatialReference = pGeoPolyline.SpatialReference
            pSplitPoint.PutCoords -180, lngSplitIndex
            pSplitPtColl.AddPoint pSplitPoint
          Next lngSplitIndex
          Set pTopoOp = pSplitPolyline
          pTopoOp.Simplify
          pSplitPolyline.SimplifyNetwork
          
          Set pTopoOp = pGeoPolyline
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
              
          Set pLeft = New Polyline
          Set pLeft.SpatialReference = pGeoPolyline.SpatialReference
          Set pRight = New Polyline
          Set pRight.SpatialReference = pGeoPolyline.SpatialReference
          pTopoOp.Cut pSplitPolyline, pLeft, pRight
          
          Set pTransform = pLeft
          pTransform.Move 360, 0
          
          Set pCombineArray = New esriSystem.varArray
          pCombineArray.Add pLeft
          pCombineArray.Add pRight
          
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pLeft, "Delete_Me"
  '
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pRight, "Delete_Me"
          
          Set pNewPolyline = MyGeometricOperations.UnionGeometries2(pCombineArray)
          Set pTopoOp = pNewPolyline
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          pNewPolyline.SimplifyNetwork
          
        
        End If
        If pEnv.XMax > 180 Then
        
          Set pSplitPolyline = New Polyline
          Set pSplitPolyline.SpatialReference = pGeoPolyline.SpatialReference
          Set pSplitPtColl = pSplitPolyline
          For lngSplitIndex = -80 To 80 Step 5
            Set pSplitPoint = New Point
            Set pSplitPoint.SpatialReference = pGeoPolyline.SpatialReference
            pSplitPoint.PutCoords 180, lngSplitIndex
            pSplitPtColl.AddPoint pSplitPoint
          Next lngSplitIndex
          Set pTopoOp = pSplitPolyline
          pTopoOp.Simplify
          pSplitPolyline.SimplifyNetwork
          
          Set pTopoOp = pGeoPolyline
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
              
          Set pLeft = New Polyline
          Set pLeft.SpatialReference = pGeoPolyline.SpatialReference
          Set pRight = New Polyline
          Set pRight.SpatialReference = pGeoPolyline.SpatialReference
          pTopoOp.Cut pSplitPolyline, pLeft, pRight
          
          Set pTransform = pRight
          pTransform.Move -360, 0
          
          Set pCombineArray = New esriSystem.varArray
  '        pCombineArray.Add pNewPolyline
          pCombineArray.Add pLeft
          pCombineArray.Add pRight
          
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pLeft, "Delete_Me"
  '
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pRight, "Delete_Me"
          
          Set pNewPolyline = MyGeometricOperations.UnionGeometries2(pCombineArray)
          Set pTopoOp = pNewPolyline
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          pNewPolyline.SimplifyNetwork
          
        End If
        
        Set SplitGeometryOnDateLine = pNewPolyline
  '      MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPolyline, "Delete_Me"
      Else
        Set SplitGeometryOnDateLine = pGeoPolyline
      End If
      
      
      
    ElseIf TypeOf pPolygonOrPolyline Is IPolygon Then
      
      Set pPolygon = pPolygonOrPolyline
      Set pClone = pPolygon
      Set pGeoPolygon = pClone.Clone
      
      Set pSpRef = pPolygon.SpatialReference
      
      If Not TypeOf pSpRef Is IGeographicCoordinateSystem Then
        Set pPrjSpRef = pSpRef
        Set pGeoSpRef = pPrjSpRef.GeographicCoordinateSystem
        pGeoPolygon.Project pGeoSpRef
      End If
      
      ' CHECK TO SEE IF THIS NEEDS TO BE SPLIT AT ALL
      Set pEnv = pGeoPolygon.Envelope
      If pEnv.XMin < -180 Or pEnv.XMax > 180 Then
   
  '      Debug.Print "Min X = " & pNewPolyline.Envelope.XMin
  '      Debug.Print "Max X = " & pNewPolyline.Envelope.XMax
              
        If pEnv.XMin < -180 Then
        
          Set pSplitPolyline = New Polyline
          Set pSplitPolyline.SpatialReference = pGeoPolygon.SpatialReference
          Set pSplitPtColl = pSplitPolyline
          For lngSplitIndex = -80 To 80 Step 5
            Set pSplitPoint = New Point
            Set pSplitPoint.SpatialReference = pGeoPolygon.SpatialReference
            pSplitPoint.PutCoords -180, lngSplitIndex
            pSplitPtColl.AddPoint pSplitPoint
          Next lngSplitIndex
          Set pTopoOp = pSplitPolyline
          pTopoOp.Simplify
          pSplitPolyline.SimplifyNetwork
                    
          Set pTopoCutter = pGeoPolygon
          Set pGeomColl = pTopoCutter.Cut2(pSplitPolyline)
          Set pCombineArray = New esriSystem.varArray
          
          For lngIndex = 0 To pGeomColl.GeometryCount - 1
            Set pSubPolygon = pGeomColl.Geometry(lngIndex)
            Set pSubEnv = pSubPolygon.Envelope
            If pSubEnv.XMin < -180 Then
              Set pTransform = pSubPolygon
              pTransform.Move 360, 0
            End If
            pCombineArray.Add pSubPolygon
          Next lngIndex
          
'
'          Set pLeft = New Polyline
'          Set pLeft.SpatialReference = pGeoPolyline.SpatialReference
'          Set pRight = New Polyline
'          Set pRight.SpatialReference = pGeoPolyline.SpatialReference
'          pTopoOp.Cut pSplitPolyline, pLeft, pRight
          
'          Set pTransform = pLeft
'          pTransform.Move 360, 0
'
'          Set pCombineArray = New esriSystem.varArray
'          pCombineArray.Add pLeft
'          pCombineArray.Add pRight
          
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pLeft, "Delete_Me"
  '
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pRight, "Delete_Me"
          
          Set pNewPolygon = MyGeometricOperations.UnionGeometries2(pCombineArray)
          Set pTopoOp = pNewPolygon
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          
        
        End If
        If pEnv.XMax > 180 Then
        
          Set pSplitPolyline = New Polyline
          Set pSplitPolyline.SpatialReference = pGeoPolygon.SpatialReference
          Set pSplitPtColl = pSplitPolyline
          For lngSplitIndex = -80 To 80 Step 5
            Set pSplitPoint = New Point
            Set pSplitPoint.SpatialReference = pGeoPolygon.SpatialReference
            pSplitPoint.PutCoords 180, lngSplitIndex
            pSplitPtColl.AddPoint pSplitPoint
          Next lngSplitIndex
          Set pTopoOp = pSplitPolyline
          pTopoOp.Simplify
          pSplitPolyline.SimplifyNetwork
                    
          Set pTopoCutter = pGeoPolygon
          Set pGeomColl = pTopoCutter.Cut2(pSplitPolyline)
          Set pCombineArray = New esriSystem.varArray
          
          For lngIndex = 0 To pGeomColl.GeometryCount - 1
            Set pSubPolygon = pGeomColl.Geometry(lngIndex)
            Set pSubEnv = pSubPolygon.Envelope
            If pSubEnv.XMax > 180 Then
              Set pTransform = pSubPolygon
              pTransform.Move -360, 0
            End If
            pCombineArray.Add pSubPolygon
          Next lngIndex
          
          
'          Set pTopoOp = pGeoPolyline
'          pTopoOp.IsKnownSimple = False
'          pTopoOp.Simplify
'
'          Set pLeft = New Polyline
'          Set pLeft.SpatialReference = pGeoPolyline.SpatialReference
'          Set pRight = New Polyline
'          Set pRight.SpatialReference = pGeoPolyline.SpatialReference
'          pTopoOp.Cut pSplitPolyline, pLeft, pRight
'
'          Set pTransform = pRight
'          pTransform.Move -360, 0
'
'          Set pCombineArray = New esriSystem.varArray
'  '        pCombineArray.Add pNewPolyline
'          pCombineArray.Add pLeft
'          pCombineArray.Add pRight
          
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pLeft, "Delete_Me"
  '
  '        MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pRight, "Delete_Me"
          
          Set pNewPolygon = MyGeometricOperations.UnionGeometries2(pCombineArray)
          Set pTopoOp = pNewPolygon
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          
        End If
        
        Set SplitGeometryOnDateLine = pNewPolygon
  '      MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me"
  '      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pNewPolyline, "Delete_Me"
      Else
        Set SplitGeometryOnDateLine = pGeoPolygon
      End If
    End If
  Else
    Set SplitGeometryOnDateLine = Nothing
  End If

  GoTo ClearMemory
ClearMemory:
  Set pGeoPolyline = Nothing
  Set pGeoPolygon = Nothing
  Set pPolyline = Nothing
  Set pPolygon = Nothing
  Set pClone = Nothing
  Set pPrjSpRef = Nothing
  Set pGeoSpRef = Nothing
  Set pSpRef = Nothing
  Set pPoint = Nothing
  Set pNewPoint = Nothing
  Set pNewPolyline = Nothing
  Set pGeomColl = Nothing
  Set pSubPolygon = Nothing
  Set pSubEnv = Nothing
  Set pNewPolygon = Nothing
  Set pSplitPolyline = Nothing
  Set pSplitPtColl = Nothing
  Set pSplitPoint = Nothing
  Set pTopoOp = Nothing
  Set pTopoCutter = Nothing
  Set pLeft = Nothing
  Set pRight = Nothing
  Set pTransform = Nothing
  Set pCombineArray = Nothing
  Set pEnv = Nothing

End Function
Public Function SpheroidalPolylineFromEndPoints2(pStartPoint As IPoint, pEndPoint As IPoint, _
    lngNumberVertices As Long)
  
  ' ASSUMES POINTS ARE IN GEOGRAPHIC COORDINATES!
  ' WILL USE DATUM OF POINTS TO GET EQUATORIAL AND POLAR RADIUS.
  
  Dim pSpRef As ISpatialReference
  Set pSpRef = pStartPoint.SpatialReference
  Dim pGeoSpRef As IGeographicCoordinateSystem
  If Not TypeOf pSpRef Is IGeographicCoordinateSystem Then
    Set SpheroidalPolylineFromEndPoints2 = Nothing
    GoTo ClearMemory
  End If
  
  Set pGeoSpRef = pSpRef
  Dim dblEquatorialRadius As Double
  Dim dblPolarRadius As Double
  dblEquatorialRadius = pGeoSpRef.Datum.Spheroid.SemiMajorAxis
  dblPolarRadius = pGeoSpRef.Datum.Spheroid.SemiMinorAxis
  
  Dim pInitialPolyline As IPointCollection
  Dim pGeom As IGeometry
  
  Set pInitialPolyline = New Polyline
  Set pGeom = pInitialPolyline
  Set pGeom.SpatialReference = pSpRef
  
  If pStartPoint.x = pEndPoint.x And pStartPoint.Y = pEndPoint.Y Then
    Set SpheroidalPolylineFromEndPoints2 = pInitialPolyline
    GoTo ClearMemory
  End If
  
  pInitialPolyline.AddPoint pStartPoint
  pInitialPolyline.AddPoint pEndPoint
  
  Dim pFinalPolyline As IPointCollection
  Set pFinalPolyline = New Polyline
  Set pGeom = pFinalPolyline
  Set pGeom.SpatialReference = pSpRef
  pFinalPolyline.AddPoint pStartPoint
  
  Dim dblIndex As Double
  Dim dblInterval As Double
  dblInterval = 1 / (lngNumberVertices - 1)
  
  Dim pPoint As IPoint
  
  For dblIndex = dblInterval To (1 - dblInterval) Step dblInterval
    Set pPoint = SpheroidalPolylineMidpoint2(pInitialPolyline, dblIndex, True, , dblEquatorialRadius, _
        dblPolarRadius)
    pFinalPolyline.AddPoint pPoint
  Next dblIndex
  
  Dim pTransform As ITransform2D
  Dim pClone As IClone
  
  If pPoint.x < -180 And pEndPoint.x > 0 Then
    Set pClone = pEndPoint
    Set pEndPoint = pClone.Clone
    Set pTransform = pEndPoint
    pTransform.Move -360, 0
  ElseIf pPoint.x > 180 And pEndPoint.x < 0 Then
    Set pClone = pEndPoint
    Set pEndPoint = pClone.Clone
    Set pTransform = pEndPoint
    pTransform.Move 360, 0
  End If
  
  pFinalPolyline.AddPoint pEndPoint
  Set SpheroidalPolylineFromEndPoints2 = pFinalPolyline
  
  GoTo ClearMemory
  
ClearMemory:
  Set pSpRef = Nothing
  Set pTransform = Nothing
  Set pClone = Nothing
  Set pGeoSpRef = Nothing
  Set pInitialPolyline = Nothing
  Set pGeom = Nothing
  Set pFinalPolyline = Nothing
  Set pPoint = Nothing

End Function

Public Function SplitMultipartFeatureIntoArray(pGeometry As IGeometry, booSucceeded As Boolean, _
    strFailureReason As String) As esriSystem.IArray
    
  strFailureReason = ""
  booSucceeded = True
  
  Dim pReturnArray As esriSystem.IArray
  Set pReturnArray = New esriSystem.Array

  Dim pPolygon As IPolygon2
  Dim pSubPolygon As IPolygon4
  Dim pPolyline As IPolyline
  Dim pMultipoint As IMultipoint
  Dim pPoint As IPoint
  Dim pPointCollection As IPointCollection
  Dim pGeometryCollection As IGeometryCollection
  Dim pOrigSegcollection As ISegmentCollection
  Dim pNewSegCollection As ISegmentCollection
  Dim pTopoOp As ITopologicalOperator2
  Dim lngNumParts As Long
  Dim pPolyComponents() As IPolygon 'Declare an array of polygon
  ReDim pPolyComponents(0)
  Dim pSpRef As ISpatialReference
  Dim lngIndex As Long
  Dim booTemp As Boolean
  
  Set pSpRef = pGeometry.SpatialReference
  
  If pGeometry.IsEmpty Then
    strFailureReason = "Empty Geometry"
    booSucceeded = False
    
  Else
    Select Case pGeometry.GeometryType
      Case esriGeometryMultipoint
        Set pMultipoint = pGeometry
        Set pPointCollection = pMultipoint
        For lngIndex = 0 To pPointCollection.PointCount - 1
          Set pPoint = pPointCollection.Point(lngIndex)
          pReturnArray.Add pPoint
        Next lngIndex
                  
      Case esriGeometryPolygon
        Set pPolygon = pGeometry
        
        ' GET CONNECTED COMPONENTS OF POLYGON
        lngNumParts = pPolygon.ExteriorRingCount
        ReDim pPolyComponents(lngNumParts - 1) 'Redimension the array of polygons with number of exterior rings
        pPolygon.GetConnectedComponents lngNumParts, pPolyComponents(0) 'Pass the first element of the array
        
'          MsgBox "Item #" & CStr(lngTimer) & vbCrLf & _
'              "Geometry Collection Count = " & CStr(pGeometryCollection.GeometryCount) & vbCrLf & _
'              "Exterior Ring Count = " & CStr(lngNumParts)
        
        For lngIndex = 0 To lngNumParts - 1
          Set pSubPolygon = pPolyComponents(lngIndex)
          Set pSubPolygon.SpatialReference = pSpRef
          Set pTopoOp = pSubPolygon
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          pReturnArray.Add pSubPolygon
        Next lngIndex
        
      Case esriGeometryPolyline
        Set pGeometryCollection = pGeometry
    
        ' GET SUB POLYLINES
        lngNumParts = pGeometryCollection.GeometryCount
        
        For lngIndex = 0 To lngNumParts - 1
          Set pOrigSegcollection = pGeometryCollection.Geometry(lngIndex)
          Set pNewSegCollection = New Polyline
          pNewSegCollection.AddSegmentCollection pOrigSegcollection
          
          Set pPolyline = pNewSegCollection
          Set pTopoOp = pPolyline
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          Set pPolyline.SpatialReference = pSpRef
          pReturnArray.Add pPolyline
          
        Next lngIndex
    End Select
  End If
  
  Set SplitMultipartFeatureIntoArray = pReturnArray
  
  GoTo ClearMemory
  
ClearMemory:
  Set pReturnArray = Nothing
  Set pPolygon = Nothing
  Set pSubPolygon = Nothing
  Set pPolyline = Nothing
  Set pMultipoint = Nothing
  Set pPoint = Nothing
  Set pPointCollection = Nothing
  Set pGeometryCollection = Nothing
  Set pOrigSegcollection = Nothing
  Set pNewSegCollection = Nothing
  Set pTopoOp = Nothing
  Erase pPolyComponents
  Set pSpRef = Nothing

End Function
Public Function SplitMultipartFeatureIntoVarArray(pGeometry As IGeometry, booSucceeded As Boolean, _
    strFailureReason As String) As Variant()
    
  strFailureReason = ""
  booSucceeded = True
  
  Dim varReturnArray() As Variant
  Dim lngCounter As Long
  
  lngCounter = -1

  Dim pPolygon As IPolygon2
  Dim pSubPolygon As IPolygon4
  Dim pPolyline As IPolyline
  Dim pMultipoint As IMultipoint
  Dim pPoint As IPoint
  Dim pPointCollection As IPointCollection
  Dim pGeometryCollection As IGeometryCollection
  Dim pOrigSegcollection As ISegmentCollection
  Dim pNewSegCollection As ISegmentCollection
  Dim pTopoOp As ITopologicalOperator2
  Dim lngNumParts As Long
  Dim pPolyComponents() As IPolygon 'Declare an array of polygon
  ReDim pPolyComponents(0)
  Dim pSpRef As ISpatialReference
  Dim lngIndex As Long
  Dim booTemp As Boolean
  
  Set pSpRef = pGeometry.SpatialReference
  
  If pGeometry.IsEmpty Then
    strFailureReason = "Empty Geometry"
    booSucceeded = False
    
  Else
    Select Case pGeometry.GeometryType
      Case esriGeometryMultipoint
        Set pMultipoint = pGeometry
        Set pPointCollection = pMultipoint
        For lngIndex = 0 To pPointCollection.PointCount - 1
          Set pPoint = pPointCollection.Point(lngIndex)
          lngCounter = lngCounter + 1
          ReDim Preserve varReturnArray(lngCounter)
          Set varReturnArray(lngCounter) = pPoint
'          pReturnArray.Add pPoint
        Next lngIndex
                  
      Case esriGeometryPolygon
        Set pPolygon = pGeometry
        
        ' GET CONNECTED COMPONENTS OF POLYGON
        lngNumParts = pPolygon.ExteriorRingCount
        ReDim pPolyComponents(lngNumParts - 1) 'Redimension the array of polygons with number of exterior rings
        pPolygon.GetConnectedComponents lngNumParts, pPolyComponents(0) 'Pass the first element of the array
        
'          MsgBox "Item #" & CStr(lngTimer) & vbCrLf & _
'              "Geometry Collection Count = " & CStr(pGeometryCollection.GeometryCount) & vbCrLf & _
'              "Exterior Ring Count = " & CStr(lngNumParts)
        
        For lngIndex = 0 To lngNumParts - 1
          Set pSubPolygon = pPolyComponents(lngIndex)
          Set pSubPolygon.SpatialReference = pSpRef
          Set pTopoOp = pSubPolygon
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          lngCounter = lngCounter + 1
          ReDim Preserve varReturnArray(lngCounter)
          Set varReturnArray(lngCounter) = pSubPolygon
'          pReturnArray.Add pSubPolygon
        Next lngIndex
        
      Case esriGeometryPolyline
        Set pGeometryCollection = pGeometry
    
        ' GET SUB POLYLINES
        lngNumParts = pGeometryCollection.GeometryCount
        
        For lngIndex = 0 To lngNumParts - 1
          Set pOrigSegcollection = pGeometryCollection.Geometry(lngIndex)
          Set pNewSegCollection = New Polyline
          pNewSegCollection.AddSegmentCollection pOrigSegcollection
          
          Set pPolyline = pNewSegCollection
          Set pTopoOp = pPolyline
          pTopoOp.IsKnownSimple = False
          pTopoOp.Simplify
          Set pPolyline.SpatialReference = pSpRef
          lngCounter = lngCounter + 1
          ReDim Preserve varReturnArray(lngCounter)
          Set varReturnArray(lngCounter) = pPolyline
'          pReturnArray.Add pPolyline
          
        Next lngIndex
    End Select
  End If
  
  SplitMultipartFeatureIntoVarArray = varReturnArray
  
  GoTo ClearMemory
  
ClearMemory:
  Erase varReturnArray
  Set pPolygon = Nothing
  Set pSubPolygon = Nothing
  Set pPolyline = Nothing
  Set pMultipoint = Nothing
  Set pPoint = Nothing
  Set pPointCollection = Nothing
  Set pGeometryCollection = Nothing
  Set pOrigSegcollection = Nothing
  Set pNewSegCollection = Nothing
  Set pTopoOp = Nothing
  Erase pPolyComponents
  Set pSpRef = Nothing



End Function


Public Function ReturnAngleOfCoverage(pOrigin As IPoint, pSinglePolylineOrPolygon As IGeometry, _
    booSucceeded As Boolean, strReason As String, Optional dblLeftmostAngle As Double, _
    Optional dblRightmostAngle As Double) As Double
  
  ' ASSUMES BOTH GEOMETRIES ARE IN THE SAME SPATIAL REFERENCE!
  
  booSucceeded = True
  strReason = ""
  Dim booIsGeographic As Boolean
  Dim dblVertices() As Double
  Dim lngIndex As Long
'  Dim pClone As IClone
'  Dim pArcGeom As IGeometry
'  Dim pArcOrigin As IPoint
  Dim dblBearing As Double
  Dim dblRight As Double
  Dim dblLeft As Double
  Dim dblOriginX As Double
  Dim dblOriginY As Double
  Dim dblPreviousCheckX As Double
  Dim dblPreviousCheckY As Double
  Dim dblCurrentCheckX As Double
  Dim dblCurrentCheckY As Double
'  Dim pSpRefFact As ISpatialReferenceFactory3
'  Dim pPrjCS As IProjectedCoordinateSystem3
'  Dim pSpRef As ISpatialReference
'  Dim booClockwise As Boolean
  Dim dblMaxRight As Double
  Dim dblMaxLeft As Double
  
  ' FOR DEBUGGING
'  Dim pMxDoc As IMxDocument
'  Dim pTestPolyline As IPolyline
'  Dim pTestPoint As IPoint
'  Dim pTestPtColl As IPointCollection
'  Set pMxDoc = ThisDocument
'  Dim dblMaxRightBearing As Double
'  Dim dblMaxLeftBearing As Double
  
  If pOrigin.IsEmpty Then
    booSucceeded = False
    strReason = "Empty Origin Point"
  ElseIf pSinglePolylineOrPolygon.IsEmpty Then
    booSucceeded = False
    strReason = "Empty Polyline or Polygon"
  ElseIf Not (TypeOf pSinglePolylineOrPolygon Is IPolygon Or TypeOf pSinglePolylineOrPolygon Is IPolyline) Then
    booSucceeded = False
    strReason = "Comparison Geometry is not polyline or polygon"
  Else
  
    booIsGeographic = TypeOf pOrigin.SpatialReference Is IGeographicCoordinateSystem
    
    ' IF ORIGIN IS IN LAT/LONG COORDINATES, THEN IT IS DIFFICULT TO DETERMINE IF CONSECUTIVE POINTS ARE CLOCKWISE
    ' OR COUNTERCLOCKWISE RELATIVE TO ORIGIN POINT.  TO SOLVE THIS, PROJECT BOTH GEOMETRIES INTO AN
    ' AZIMUTHAL EQUIDISTANT PROJECTION CENTERED ON THE ORIGIN.  THE RELATIVE DIRECTIONS SHOULD STILL BE THE SAME.
    
    ' NEW PLAN.  DON'T WORRY ABOUT CONSECUTIVE POINTS CLOCKWISE FROM COORDINATES; INSTEAD BASE EVERYTHING ON
    ' BEARINGS.  THIS TIME WE HAVE TO ASSUME BOTH GEOMETRIES ARE IN THE SAME SPATIAL REFERENCE.
    
    
    
'    Set pClone = pOrigin
'    Set pArcOrigin = pClone.Clone
    
'    If TypeOf pArcOrigin.SpatialReference Is IGeographicCoordinateSystem Then
'
'       ' PROJECT INTO AZIMUTHAL EQUIDISTANT
'      Set pSpRefFact = New SpatialReferenceEnvironment
'      Set pPrjCS = pSpRefFact.CreateProjectedCoordinateSystem(esriSRProjCS_World_AzimuthalEquidistant)
'      Set pSpRef = pPrjCS
'      pPrjCS.CentralMeridian(True) = pArcOrigin.X
'      pPrjCS.LatitudeOfOrigin = pArcOrigin.Y
'
'      If Not MyGeomCheckSpRefDomain(pSpRef) Then
'        Dim pSpRefRes As ISpatialReferenceResolution
'        Set pSpRefRes = pSpRef
'        pSpRefRes.ConstructFromHorizon
'      End If
'      pArcOrigin.Project pPrjCS
'    End If
    
    ' MAKE SURE POLYGON/POLYLINE IS IN SAME COORDINATE SYSTEM
'    Set pClone = pSinglePolylineOrPolygon
'    Set pArcGeom = pClone.Clone
'    If Not MyGeneralOperations.CompareSpatialReferences(pArcOrigin.SpatialReference, _
'        pArcGeom.SpatialReference) Then
'      pArcGeom.Project pArcOrigin.SpatialReference
'    End If
    
    
    ' FOR DEBUGGING
'    MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pArcGeom, "Delete_me"
    
    ' CONVERT POLYGON/POLYLINE TO VERTICES FOR FASTER PROCESSING
'    dblVertices = MyGeometricOperations.ReturnVerticesAsDoubleArray(pArcGeom)
    dblVertices = MyGeometricOperations.ReturnVerticesAsDoubleArray(pSinglePolylineOrPolygon)
    
'    dblOriginX = pArcOrigin.X
'    dblOriginY = pArcOrigin.Y
    dblOriginX = pOrigin.x
    dblOriginY = pOrigin.Y
    
    Dim dblOriginalBearing As Double
    Dim dblCurrentSweepPosition As Double
    Dim dblPreviousBearing As Double
    Dim dblCurrentDeviation As Double
    Dim dblDistance As Double
    Dim dblAZ2 As Double
    
    ' FOR EACH VERTEX, CALCULATE BEARING AND WEHTHER THAT BEARING IS CLOCKWISE OR COUNTERCLOCKWISE TO THE PREVIOUS BEARING
    For lngIndex = 0 To UBound(dblVertices, 2)
      
      dblCurrentCheckX = dblVertices(0, lngIndex)
      dblCurrentCheckY = dblVertices(1, lngIndex)
      
      If booIsGeographic Then
'        dblHavDist = DistanceHaversineNumbers(dblOriginY, dblOriginX, dblCurrentCheckY, dblCurrentCheckX, _
            , True, dblBearing)
        dblDistance = DistanceVincentyNumbers2(dblOriginX, dblOriginY, dblCurrentCheckX, dblCurrentCheckY, dblBearing, _
            dblAZ2)
      Else
        dblBearing = CalcBearingNumbers(dblOriginX, dblOriginY, dblCurrentCheckX, dblCurrentCheckY)
      End If
      
      If lngIndex = 0 Then
        dblMaxLeft = 0
        dblMaxRight = 0
        dblOriginalBearing = dblBearing
        dblCurrentDeviation = 0
      Else
        ' CHECK IF THIS VERTEX APPEARS CLOCWISE OR COUNTERCLOCKWISE FROM PREVIOUS BEARING, RELATIVE TO ORIGIN
'        booClockwise = CalcCheckClockwiseNumbers(dblOriginX, dblOriginY, dblPreviousCheckX, _
            dblPreviousCheckY, dblCurrentCheckX, dblCurrentCheckY)
            
        dblCurrentDeviation = CalcDirectionDeviationDegrees(dblPreviousBearing, _
            dblBearing)
        dblCurrentSweepPosition = dblCurrentSweepPosition + dblCurrentDeviation
        If dblCurrentSweepPosition < dblMaxLeft Then dblMaxLeft = dblCurrentSweepPosition
        If dblCurrentSweepPosition > dblMaxRight Then dblMaxRight = dblCurrentSweepPosition
               
            
'        If booClockwise Then
'          ' CHECK IF CURRENT BEARING IS FARTHER RIGHT THAN THE PREVIOUS MAXIMUM RIGHT
'          dblRight = CalcDirectionDeviationDegrees(dblMaxRight, dblBearing)
'          If dblRight > 0 Then dblMaxRight = dblBearing
'        Else
'          ' CHECK IF CURRENT BEARING IS FARTHER LEFT THAN THE PREVIOUS MAXIMUM LEFT
'          dblLeft = CalcDirectionDeviationDegrees(dblMaxLeft, dblBearing)
'          If dblLeft < 0 Then dblMaxRight = dblBearing
'        End If
          
      End If
      
      
      ' FOR DEBUGGING
'      If lngIndex > 380 Then
'        Set pTestPoint = New Point
'        Set pTestPoint.SpatialReference = pArcOrigin.SpatialReference
'        pTestPoint.PutCoords dblOriginX, dblOriginY
'        Set pTestPolyline = New Polyline
'        Set pTestPolyline.SpatialReference = pArcOrigin.SpatialReference
'        Set pTestPtColl = pTestPolyline
'        pTestPtColl.AddPoint pTestPoint
'        Set pTestPoint = New Point
'        Set pTestPoint.SpatialReference = pArcOrigin.SpatialReference
'        pTestPoint.PutCoords dblCurrentCheckX, dblCurrentCheckY
'        pTestPtColl.AddPoint pTestPoint
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pTestPolyline, "Delete_me"
'        Debug.Print "Current Bearing = " & Format(dblBearing, "0.00")
'        dblMaxLeftBearing = dblOriginalBearing + dblMaxLeft
'        dblMaxRightBearing = dblOriginalBearing + dblMaxRight
'        ForceAzimuthToCorrectRange dblMaxLeftBearing
'        ForceAzimuthToCorrectRange dblMaxRightBearing
'        Debug.Print "  --> Current Maximum Left Bearing = " & Format(dblMaxLeftBearing, "0.00")
'        Debug.Print "  --> Current Maximum Right Bearing = " & Format(dblMaxRightBearing, "0.00")
'        Debug.Print "  --> Current Maximum Left Deviation = " & Format(dblMaxLeft, "0.00")
'        Debug.Print "  --> Current Maximum Right Deviation = " & Format(dblMaxRight, "0.00")
'        Debug.Print "  --> Current Sweep Position = " & Format(dblCurrentSweepPosition, "0.00")
'  '      Debug.Print "  --> CW Deviation of Current Bearing from Maximum Right Bearing = " & Format(dblRight, "0.00")
'  '      Debug.Print "  --> CCW Deviation of Current Bearing from Maximum Left Bearing = " & Format(dblLeft, "0.00")
'        Debug.Print "  --> Current Bearing Clocwise from Previous Bearing = " & UCase(CStr(booClockwise))
'      End If
      
      
      dblPreviousCheckX = dblCurrentCheckX
      dblPreviousCheckY = dblCurrentCheckY
      dblPreviousBearing = dblBearing
      
      
    Next lngIndex
  End If
  
  ReturnAngleOfCoverage = Abs(dblMaxLeft) + Abs(dblMaxRight)
  dblLeftmostAngle = dblOriginalBearing + dblMaxLeft   ' dblMaxLeft will be negative
  dblRightmostAngle = dblOriginalBearing + dblMaxRight
  ForceAzimuthToCorrectRange dblLeftmostAngle
  ForceAzimuthToCorrectRange dblRightmostAngle
  
  
  
  GoTo ClearMemory
ClearMemory:
  Erase dblVertices

End Function


Public Function ConvertRotationMathRadiansToCompassDegrees(dblRadiansCCW As Double) As Double

  Dim dblDegrees As Double
  dblDegrees = -RadToDeg(dblRadiansCCW)
  
  ConvertRotationMathRadiansToCompassDegrees = dblDegrees

End Function


Public Function ConvertRotationCompassDegreesToMathRadians(dblCompassClockwise As Double) As Double

  Dim dblRadians As Double
  dblRadians = -DegToRad(dblCompassClockwise)
  
  ConvertRotationCompassDegreesToMathRadians = dblRadians

End Function

Public Function ReturnConvexHullFromFClass(pFLayer As IFeatureLayer, _
    Optional booUseCurrentlySelected As Boolean = False, Optional booMakeNewSelection As Boolean = False, _
    Optional strQueryString As String) As IPolygon

'  'SAMPLE CODE
'  Dim dblArea As Double
'  Dim strQuery As String
'  Dim pFLayer As IFeatureLayer
'  Dim pMxDoc As IMxDocument
'  Dim strPrefix As String
'  Dim strSuffix As String
'  Dim pPolygon As IPolygon
'  Dim pArea As IArea
'
'  Set pMxDoc = ThisDocument
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("Owl_83_Day_Night", pMxDoc.FocusMap)
'  MyGeneralOperations.ReturnQuerySpecialCharacters pFLayer.FeatureClass, strPrefix, strSuffix
'  strQuery = strPrefix & "Period" & strSuffix & " = 'Day'"
'  Set pPolygon = ReturnConvexHullFromFClass(pFLayer, False, True, strQuery)
'  Set pArea = pPolygon
'  Debug.Print "Acreage = " & Format(pArea.Area / 4046.8564224, "#,##0.000")
'
'ClearMemory:
'  Set pFLayer = Nothing
'  Set pMxDoc = Nothing
'  Set pPolygon = Nothing
'  Set pArea = Nothing

  Dim pGeomBag As IGeometryBag
  Dim pFClass As IFeatureClass
  Dim pFCursor As IFeatureCursor
  Dim pGeomArray As IArray
  Dim pFeatSel As IFeatureSelection
  Dim pSelSet As ISelectionSet
  Dim pGeoDataset As IGeoDataset
  Dim pFeature As IFeature
  Dim pGeom As IGeometry
  Dim pQueryFilt As IQueryFilter
  
  Dim pPtColl As IPointCollection
  Set pPtColl = New Multipoint
  Dim pGeomPtColl As IPointCollection
  Dim pPoint As IPoint
  
  Set pGeomArray = New esriSystem.Array
  Set pFClass = pFLayer.FeatureClass
  Set pGeoDataset = pFClass
  
  If booUseCurrentlySelected Then
    Set pFeatSel = pFLayer
    Set pSelSet = pFeatSel.SelectionSet
    If pSelSet.Count = 0 Then
      Set ReturnConvexHullFromFClass = New Polygon  ' RETURN AN EMPTY POLYGON
      Set ReturnConvexHullFromFClass.SpatialReference = pGeoDataset.SpatialReference
      GoTo ClearMemory
    Else
      pSelSet.Search Nothing, False, pFCursor
    End If
  ElseIf booMakeNewSelection Then
    Set pQueryFilt = New QueryFilter
    pQueryFilt.WhereClause = strQueryString
    Set pFCursor = pFClass.Search(pQueryFilt, False)
  Else
    Set pFCursor = pFClass.Search(Nothing, False)
  End If
    
  Set pFeature = pFCursor.NextFeature
  Do Until pFeature Is Nothing
    
    Set pGeom = pFeature.ShapeCopy
    
    If Not pGeom.IsEmpty Then
      If TypeOf pGeom Is IPoint Then
        pPtColl.AddPoint pGeom
      Else
        Set pGeomPtColl = pGeom
        pPtColl.AddPointCollection pGeomPtColl
      End If
    End If
    Set pFeature = pFCursor.NextFeature
  Loop

  If pPtColl.PointCount = 0 Then
    Set ReturnConvexHullFromFClass = New Polygon  ' RETURN AN EMPTY POLYGON
    Set ReturnConvexHullFromFClass.SpatialReference = pGeoDataset.SpatialReference
    GoTo ClearMemory
  Else
    Set ReturnConvexHullFromFClass = ReturnConvexHullFromGeometry(pPtColl)
  End If
  
  GoTo ClearMemory
  
ClearMemory:
  Set pGeomBag = Nothing
  Set pFClass = Nothing
  Set pFCursor = Nothing
  Set pGeomArray = Nothing
  Set pFeatSel = Nothing
  Set pSelSet = Nothing
  Set pGeoDataset = Nothing
  Set pFeature = Nothing
  Set pGeom = Nothing
  Set pPtColl = Nothing
  Set pGeomPtColl = Nothing
  Set pQueryFilt = Nothing

End Function


Public Function ReturnConvexHullFromGeometry(pGeom As IGeometry) As IPolygon

  ' IF pGeom IS A POINT, THEN THE HULL IS ALSO A POINT!
  
  Dim pTopoOp As ITopologicalOperator
  Dim pPoint As IPoint
  Dim pHull_1 As IGeometry
  Dim pEnv As IEnvelope
  Dim pHull_2 As IPolygon
  Dim pPolyline As IPolyline
  Dim pPtColl As IPointCollection
  
  Set pTopoOp = pGeom
  Set pHull_1 = pTopoOp.ConvexHull
  
  If TypeOf pHull_1 Is IPoint Then
    Set pPoint = pHull_1
    If pPoint.IsEmpty Then
      Set pHull_2 = New Polygon
      Set pHull_2.SpatialReference = pGeom.SpatialReference
    Else
      Set pEnv = pPoint.Envelope
      Set pHull_2 = EnvelopeToPolygon(pEnv)
    End If
  ElseIf TypeOf pHull_1 Is IPolyline Then
    Set pPtColl = New Polygon
    Set pPolyline = pHull_1
    pPtColl.AddPoint pPolyline.FromPoint
    pPtColl.AddPoint pPolyline.ToPoint
    Set pHull_2 = pPtColl
    pHull_2.Close
    Set pHull_2.SpatialReference = pGeom.SpatialReference
  Else
    Set pHull_2 = pHull_1
  End If
  
  Set ReturnConvexHullFromGeometry = pHull_2
  
  GoTo ClearMemory
ClearMemory:
  Set pTopoOp = Nothing
  Set pPoint = Nothing
  Set pHull_1 = Nothing
  Set pEnv = Nothing
  Set pHull_2 = Nothing
  Set pPolyline = Nothing
  Set pPtColl = Nothing
  
End Function


Public Function ConvertAngleMathRadiansToCompassDegrees(dblRadiansCCW As Double) As Double
  
  ' ACCOUNTS FOR DIFFERENT 0-POINT VALUES BETWEEN MATH AND COMPASS DIRECTION
  
  Dim dblDegrees As Double
  dblDegrees = -RadToDeg(dblRadiansCCW) + 90
  
  ConvertAngleMathRadiansToCompassDegrees = dblDegrees

End Function


Public Function ConvertAngleCompassDegreesToMathRadians(dblCompassClockwise As Double) As Double

  ' ACCOUNTS FOR DIFFERENT 0-POINT VALUES BETWEEN MATH AND COMPASS DIRECTION
  
  Dim dblRadians As Double
  dblRadians = -DegToRad(dblCompassClockwise - 90)
  
  ConvertAngleCompassDegreesToMathRadians = dblRadians

End Function

Public Function ReturnDecimalMagnitude(dblVal As Double, Optional booFound As Boolean) As Long
  
  booFound = False
  Dim dblExp As Double
  Dim dblTest As Double
  dblTest = Abs(dblVal)
  
  For dblExp = -323 To 308 Step 1
    If 10 ^ dblExp > dblTest Then
      ReturnDecimalMagnitude = dblExp
      booFound = True
      Exit For
    End If
  Next dblExp

End Function




Public Function LogX(dblBase As Double, dblVal As Double) As Double

   LogX = Log(dblVal) / Log(dblBase)

End Function

Public Function ReturnDecimalMagnitude2(dblVal As Double, Optional booFound As Boolean) As Long
  ' returns 0 for [0 - 9.999], 1 for [10 - 19.999], etc.
    
  ReturnDecimalMagnitude2 = Int(LogX(10, dblVal))
    
End Function

Function Ceiling(ByVal num As Double) As Long

    Dim x As Long

    x = Int(num)
    Ceiling = x + IIf(x = num, 0#, 1#)

End Function

Function MinLong(lngX As Long, lngY As Long) As Long

  If lngX < lngY Then
    MinLong = lngX
  Else
    MinLong = lngY
  End If

End Function

Function MaxLong(lngX As Long, lngY As Long) As Long

  If lngX > lngY Then
    MaxLong = lngX
  Else
    MaxLong = lngY
  End If

End Function

Function MinDouble(dblX As Double, dblY As Double) As Double

  If dblX < dblY Then
    MinDouble = dblX
  Else
    MinDouble = dblY
  End If

End Function

Function MaxDouble(dblX As Double, dblY As Double) As Double

  If dblX > dblY Then
    MaxDouble = dblX
  Else
    MaxDouble = dblY
  End If

End Function

Public Function NiceNumber(dblX As Double, booRound As Boolean) As Double

  ' ADAPTED FROM "GRAPHIC GEMS" BY ANDREW S. GLASSNER (ACADEMIC PRESS, 1993), P. 61-63 ["NICE NUMBERS FOR GRAPH LABELS"]
  ' Returns a "nice" number approximately equal to dblX.  Rounds the number of booRound = True, otherwise takes the ceiling of the number
  
  Dim lngExp As Long
  Dim dblFraction As Double
  Dim dblRoundFrac As Double
  
  lngExp = Int(LogX(10, dblX))             ' GETS THE MAGNITUDE OF THE NUMBER
  dblFraction = dblX / (10 ^ lngExp)       ' BETWEEN 1 AND 10
  
  If booRound Then
    If dblFraction < 1.5 Then
      dblRoundFrac = 1
    ElseIf dblFraction < 3 Then
      dblRoundFrac = 2
    ElseIf dblFraction < 7 Then
      dblRoundFrac = 5
    Else
      dblRoundFrac = 10
    End If
  Else
    If dblFraction <= 1 Then
      dblRoundFrac = 1
    ElseIf dblFraction <= 2 Then
      dblRoundFrac = 2
    ElseIf dblFraction <= 5 Then
      dblRoundFrac = 5
    Else
      dblRoundFrac = 10
    End If
  End If
    
  NiceNumber = dblRoundFrac * (10 ^ lngExp)

End Function

Public Function ReturnRoundedIntervals2(dblMinimum As Double, dblMaximum As Double, lngMinIntervals As Long, _
    strTextValuesToFill() As String, Optional dblIntervalToFill As Double, Optional dblGraphMinToFill As Double, _
    Optional dblGraphMaxToFill As Double, Optional dblConversionFactor As Double = 1, _
    Optional strFormatStringToFill As String = "0", Optional booSucceeded As Boolean, _
    Optional dblConvertedMinVal As Double, Optional strConvertedMinText As String, _
    Optional dblConvertedMaxVal As Double, Optional strConvertedMaxText As String, _
    Optional dblConvertedIntervalVal As Double, Optional strConvertedIntervalText As String) As Double()
    

  ' ADAPTED FROM "GRAPHIC GEMS" BY ANDREW S. GLASSNER (ACADEMIC PRESS, 1993), P. 61-63 ["NICE NUMBERS FOR GRAPH LABELS"]
  ' dblGraphMinToFill, dblGraphMaxToFill, dblIntervalToFill and all Tic numeric values are in unconverted units
  ' All strTextValuesToFill() values are in converted units
   
  Dim dblConvertMaximum As Double
  Dim dblConvertMinimum As Double
  Dim dblTemp As Double
  dblConvertMaximum = dblMaximum * dblConversionFactor
  dblConvertMinimum = dblMinimum * dblConversionFactor
  
  Dim dblReturn() As Double
  If dblConvertMaximum = dblConvertMinimum Then
    booSucceeded = False
    GoTo ClearMemory
  ElseIf dblConvertMaximum < dblConvertMinimum Then
    dblTemp = dblConvertMaximum
    dblConvertMaximum = dblConvertMinimum
    dblConvertMinimum = dblTemp
  End If
  
  Dim intNFrac As Integer
'  Dim dblD As Double
  Dim dblGraphMin As Double
  Dim dblGraphMax As Double
  Dim dblRange As Double
  Dim dblX As Double
  Dim dblTempGraphMin As Double
  Dim dblTempGraphMax As Double
  
 ' MsgBox "In ReturnRoundedIntervals2:" & vbCrLf & _
      "dblConvertMinimum = " & CStr(dblConvertMinimum) & vbCrLf & _
      "dblConvertMaximum = " & CStr(dblConvertMaximum) & vbCrLf & _
      "dblConversionFactor = " & CStr(dblConversionFactor)

  
  dblRange = NiceNumber(dblConvertMaximum - dblConvertMinimum, False)
  dblIntervalToFill = NiceNumber(dblRange / CDbl(lngMinIntervals - 1), True)
  dblTempGraphMin = CDbl(Int(dblConvertMinimum / dblIntervalToFill)) * dblIntervalToFill
  dblTempGraphMax = CDbl(Ceiling(dblConvertMaximum / dblIntervalToFill)) * dblIntervalToFill
  intNFrac = MaxLong(-Int(LogX(10, dblIntervalToFill)), 0)
  
  If intNFrac = 0 Then
    strFormatStringToFill = "0"
  Else
    strFormatStringToFill = "0." & String(intNFrac, "0")
  End If
    
  Dim lngCounter As Long
  lngCounter = -1
  
  Dim dblInterval As Double
  For dblInterval = dblTempGraphMin To dblTempGraphMax + (dblIntervalToFill / 2) Step dblIntervalToFill
    lngCounter = lngCounter + 1
    dblConvertedMaxVal = dblInterval
    
    ReDim Preserve dblReturn(lngCounter)
    dblReturn(lngCounter) = dblInterval / dblConversionFactor
    ReDim Preserve strTextValuesToFill(lngCounter)
    strTextValuesToFill(lngCounter) = Format(dblInterval, strFormatStringToFill)
  Next dblInterval
      
  dblConvertedIntervalVal = dblIntervalToFill
  dblConvertedMinVal = dblTempGraphMin
  
  strConvertedIntervalText = Format(dblConvertedIntervalVal, strFormatStringToFill)
  strConvertedMinText = Format(dblConvertedMinVal, strFormatStringToFill)
  strConvertedMaxText = Format(dblConvertedMaxVal, strFormatStringToFill)
  
  dblIntervalToFill = dblIntervalToFill / dblConversionFactor
  dblGraphMinToFill = dblReturn(0)
  dblGraphMaxToFill = dblReturn(UBound(dblReturn))
  
  ReturnRoundedIntervals2 = dblReturn
  
  GoTo ClearMemory
ClearMemory:
  Erase dblReturn
End Function


Public Function ReturnPolygonRingsAsDoubleArray(pPolygon As IPolygon) As Variant()
    
  ' RETURNS A VARIANT ARRAY CONTAINING MULTIPLE 2-DIMENSIONAL DOUBLE ARRAYS OF X- AND Y-COORDINATES OF ALL POLYGON RINGS.
  ' EACH DOUBLE ARRAY HAS (VERTEX COUNT) ROWS, WHERE THE LAST ROW/VERTEX IS EQUAL TO THE FIRST ROW/VERTEX.
  ' CODE REMOVES ANY DUPLICATE CONSECUTIVE VERTICES.
  
  Dim varReturn() As Variant
  Dim pGeomColl As IGeometryCollection
  Dim pRing As IRing
  Dim lngIndex As Long
  Dim lngCounter As Long
  Dim lngRingCounter As Long
  Dim pPtColl As IPointCollection
  Dim pTestPoint1 As IPoint
  Dim lngIndex2 As Long
  Dim lngPointCount As Long
  Dim pPoint As IPoint
  Dim dblArray() As Double
  Dim booAdd As Boolean
  Dim pClone As IClone
  Dim pWorkPolygon As IPolygon
  
  If pPolygon.IsEmpty Then
    ReDim varReturn(0)
    varReturn(0) = Null
  Else
  
    Set pPoint = New Point
    Set pClone = pPolygon
    Set pWorkPolygon = pClone.Clone
    pWorkPolygon.SimplifyPreserveFromTo
    
    lngRingCounter = -1
    Set pGeomColl = pWorkPolygon
    For lngIndex = 0 To pGeomColl.GeometryCount - 1
      Set pRing = pGeomColl.Geometry(lngIndex)
      If Not pRing.IsEmpty Then
        If Not pRing.IsClosed Then pRing.Close
        Set pPtColl = pRing
        lngCounter = -1
        For lngIndex2 = 0 To pPtColl.PointCount - 1
          pPtColl.QueryPoint lngIndex2, pPoint
          booAdd = True
          If lngCounter > -1 Then
            If pPoint.x = dblArray(0, lngCounter) And pPoint.Y = dblArray(1, lngCounter) Then booAdd = False
          End If
          
          If booAdd Then
            lngCounter = lngCounter + 1
            ReDim Preserve dblArray(1, lngCounter)
            dblArray(0, lngCounter) = pPoint.x
            dblArray(1, lngCounter) = pPoint.Y
          End If
        Next lngIndex2
        
        ' BELOW SHOULD NEVER HAPPEN, BUT FORCE POLYGON CLOSE IF NECESSARY
        If lngCounter > -1 And ((dblArray(0, 0) <> dblArray(0, lngCounter)) Or (dblArray(0, 0) <> dblArray(0, lngCounter))) Then
          lngCounter = lngCounter + 1
          ReDim Preserve dblArray(1, lngCounter)
          dblArray(0, lngCounter) = dblArray(0, 0)
          dblArray(1, lngCounter) = dblArray(1, 0)
        End If
        
        lngRingCounter = lngRingCounter + 1
        ReDim Preserve varReturn(lngRingCounter)
        varReturn(lngRingCounter) = dblArray
      End If
    Next lngIndex
  End If
  
  ReturnPolygonRingsAsDoubleArray = varReturn
        
End Function
Public Sub TestIntersect()
 

  Dim pPoint1 As IPoint
  Dim pPoint2 As IPoint
  Set pPoint1 = New Point
  Set pPoint2 = New Point
  pPoint1.PutCoords 100, 400
  pPoint2.PutCoords 400, 100
 
  Dim pPolygon1 As IPolygon
  Dim pPolygon2 As IPolygon
 
  Set pPolygon1 = MyGeometricOperations.CreateCircleAroundPoint(pPoint1, 100, 360)
  Set pPolygon2 = MyGeometricOperations.CreateCircleAroundPoint(pPoint2, 100, 360)
  Dim pEnv1 As IEnvelope
  Dim pEnv2 As IEnvelope
  Set pEnv1 = pPolygon1.Envelope
  Set pEnv2 = pPolygon2.Envelope
  Debug.Print CStr(EnvelopeIntersects(pEnv1.XMin, pEnv1.XMax, pEnv1.YMin, pEnv1.YMax, pEnv2.XMin, _
      pEnv2.XMax, pEnv2.YMin, pEnv2.YMax))
     
ClearMemory:
  Set pPoint1 = Nothing
  Set pPoint2 = Nothing
  Set pPolygon1 = Nothing
  Set pPolygon2 = Nothing
  Set pEnv1 = Nothing
  Set pEnv2 = Nothing
 
End Sub

 

Public Function EnvelopeIntersects(dblXMin_A As Double, dblXMax_A As Double, dblYMin_A As Double, _
    dblYMax_A As Double, dblXMin_B As Double, dblXMax_B As Double, dblYMin_B As Double, _
    dblYMax_B As Double) As Boolean
   
'  Dim pPoint1 As IPoint
'  Dim pPoint2 As IPoint
'  Set pPoint1 = New Point
'  Set pPoint2 = New Point
'  pPoint1.PutCoords 100, 100
'  pPoint2.PutCoords 200, 100
'
'  Dim pPolygon1 As IPolygon
'  Dim pPolygon2 As IPolygon
'
'  Set pPolygon1 = MyGeometricOperations.CreateCircleAroundPoint(pPoint1, 100, 360)
'  Set pPolygon2 = MyGeometricOperations.CreateCircleAroundPoint(pPoint2, 100, 360)
'  Dim pEnv1 As IEnvelope
'  Dim pEnv2 As IEnvelope
'  Set pEnv1 = pPolygon1.Envelope
'  Set pEnv2 = pPolygon2.Envelope
'  Debug.Print CStr(EnvelopeIntersects(pEnv1.XMin, pEnv1.XMax, pEnv1.YMin, pEnv1.YMax, pEnv2.XMin, _
'      pEnv2.XMax, pEnv2.YMin, pEnv2.YMax))
'
'ClearMemory:
'  Set pPoint1 = Nothing
'  Set pPoint2 = Nothing
'  Set pPolygon1 = Nothing
'  Set pPolygon2 = Nothing
'  Set pEnv1 = Nothing
'  Set pEnv2 = Nothing
  EnvelopeIntersects = Not (dblXMin_A > dblXMax_B Or dblYMin_A > dblYMax_B Or dblXMax_A < dblXMin_B Or dblYMax_A < dblYMin_B)
 
End Function
 
Public Function PointInPolygon(dblX As Double, dblY As Double, dblPolygon() As Double) As Boolean
' adapted from http://geomalgorithms.com/a03-_inclusion.html
'// Copyright 2000 softSurfer, 2012 Dan Sunday
'// This code may be freely used and modified for any purpose
'// providing that this copyright notice is included with it.
'// SoftSurfer makes no warranty for this code, and cannot be held
'// liable for any real or imagined damage resulting from its use.
'// Users of this code must verify correctness for their application.
'
'
'// a Point is defined by its coordinates {int x, y;}
'//===================================================================
'
'
'// isLeft(): tests if a point is Left|On|Right of an infinite line.
'//    Input:  three points P0, P1, and P2
'//    Return: >0 for P2 left of the line through P0 and P1
'//            =0 for P2  on the line
'//            <0 for P2  right of the line
'//    See: Algorithm 1 "Area of Triangles and Polygons"
'inline int
'isLeft( Point P0, Point P1, Point P2 )
'{
'    return ( (P1.x - P0.x) * (P2.y - P0.y)
'            - (P2.x -  P0.x) * (P1.y - P0.y) );
'}
'//===================================================================
'
'
'// cn_PnPoly(): crossing number test for a point in a polygon
'//      Input:   P = a point,
'//               V[] = vertex points of a polygon V[n+1] with V[n]=V[0]
'//      Return:  0 = outside, 1 = inside
'// This code is patterned after [Franklin, 2000]
'int
'cn_PnPoly( Point P, Point* V, int n )'{
'    int    cn = 0;    // the  crossing number counter
'
'    // loop through all edges of the polygon
'    for (int i=0; i<n; i++) {    // edge from V[i]  to V[i+1]
'       if (((V[i].y <= P.y) && (V[i+1].y > P.y))     // an upward crossing
'        || ((V[i].y > P.y) && (V[i+1].y <=  P.y))) { // a downward crossing
'            // compute  the actual edge-ray intersect x-coordinate
'            float vt = (float)(P.y  - V[i].y) / (V[i+1].y - V[i].y);
'            if (P.x <  V[i].x + vt * (V[i+1].x - V[i].x)) // P.x < intersect
'                 ++cn;   // a valid crossing of y=P.y right of P.x
'        }
'    }
'    return (cn&1);    // 0 if even (out), and 1 if  odd (in)
'
'}
'//===================================================================
'
'
'// wn_PnPoly(): winding number test for a point in a polygon
'//      Input:   P = a point,
'//               V[] = vertex points of a polygon V[n+1] with V[n]=V[0]
'//      Return:  wn = the winding number (=0 only when P is outside)
'int
'wn_PnPoly( Point P, Point* V, int n )
'{
'    int    wn = 0;    // the  winding number counter
'
'    // loop through all edges of the polygon
'    for (int i=0; i<n; i++) {   // edge from V[i] to  V[i+1]
'        if (V[i].y <= P.y) {          // start y <= P.y
'            if (V[i+1].y  > P.y)      // an upward crossing
'                 if (isLeft( V[i], V[i+1], P) > 0)  // P left of  edge
'                     ++wn;            // have  a valid up intersect
'        }
'        else {                        // start y > P.y (no test needed)
'            if (V[i+1].y  <= P.y)     // a downward crossing
'                 if (isLeft( V[i], V[i+1], P) < 0)  // P right of  edge
'                     --wn;            // have  a valid down intersect
'        }
'    }
'    return wn;
'}
'//===================================================================
   
End Function

Public Function PointInPoly_Winding(dblPointX As Double, dblPointY As Double, varPolygon() As Variant, _
    Optional lngWindCounter As Long) As Boolean

  ' ASSUMES POLYGON IS IN THE FORM OF A VARIANT ARRAY, WHERE EACH OBJECT IN THE ARRAY IS A POLYGON RING.
  ' EACH RING IS IN THE FORM OF A DOUBLE-ARRAY OF EACH VERTEX IN THE RING.
  ' VERTEX (0) = VERTEX (uBound(RingArray))
  '
  ' USES FUNCTION CalcCheckClockwiseNumbers2, BUT HARD-CODED DIRECTLY FOR PERFORMANCE
  '
  ' IN TESTS, WINDING FUNCTION GIVES SAME RESULTS AS CROSS-METHOD, BUT RUNS ROUGHLY 9% FASTER
  ' IN GENERAL BOTH WINDING AND CROSS METHODS GIVE ACCURATE RESULTS IN MULTPART POLYGONS CONTAINING ISLANDS AND NESTED HOLES
  
  ' adapted from http://geomalgorithms.com/a03-_inclusion.html
  '// isLeft(): tests if a point is Left|On|Right of an infinite line.
  '//    Input:  three points P0, P1, and P2
  '//    Return: >0 for P2 left of the line through P0 and P1
  '//            =0 for P2  on the line
  '//            <0 for P2  right of the line
  '//    See: Algorithm 1 "Area of Triangles and Polygons"
  'inline int
  'isLeft( Point P0, Point P1, Point P2 )
  '{
  '    return ( (P1.x - P0.x) * (P2.y - P0.y)
  '            - (P2.x -  P0.x) * (P1.y - P0.y) );
  '}
  '// wn_PnPoly(): winding number test for a point in a polygon
  '//      Input:   P = a point,
  '//               V[] = vertex points of a polygon V[n+1] with V[n]=V[0]
  '//      Return:  wn = the winding number (=0 only when P is outside)
  'int
  'wn_PnPoly( Point P, Point* V, int n )
  '{
  '    int    wn = 0;    // the  winding number counter
  '
  '    // loop through all edges of the polygon
  '    for (int i=0; i<n; i++) {   // edge from V[i] to  V[i+1]
  '        if (V[i].y <= P.y) {          // start y <= P.y
  '            if (V[i+1].y  > P.y)      // an upward crossing
  '                 if (isLeft( V[i], V[i+1], P) > 0)  // P left of  edge
  '                     ++wn;            // have  a valid up intersect
  '        }
  '        else {                        // start y > P.y (no test needed)
  '            if (V[i+1].y  <= P.y)     // a downward crossing
  '                 if (isLeft( V[i], V[i+1], P) < 0)  // P right of  edge
  '                     --wn;            // have  a valid down intersect
  '        }
  '    }
  '    return wn;
  '}
  '//===================================================================

  Dim lngRingIndex As Long
  Dim dblRing() As Double
  Dim lngPointIndex As Long
  
  ' TEST IF THIS IS FASTER
  ' YES; ROUGHLY 20% FASTER IF I DECLARE THESE VARIABLES RATHER THAN REFERENCE THE ARRAY DIRECLTY
  Dim dblX1 As Double
  Dim dblY1 As Double
  Dim dblX2 As Double
  Dim dblY2 As Double
  Dim dblSegProportion As Double
    
  lngWindCounter = 0
  
  For lngRingIndex = 0 To UBound(varPolygon)
    dblRing = varPolygon(lngRingIndex)
    dblX1 = dblRing(0, 0)
    dblY1 = dblRing(1, 0)
    
    For lngPointIndex = 0 To UBound(dblRing, 2) - 1
      ' This "if-then" excludes all cases where the edge segment can't possibly intersect horizontal line, and
      ' also cases where edge segment itself is horizontal
      dblX2 = dblRing(0, lngPointIndex + 1)
      dblY2 = dblRing(1, lngPointIndex + 1)
      
      If dblY1 <= dblPointY Then
        If dblY2 > dblPointY Then  ' If an upward Crossing
'          If ((dblX2 * (dblPointY - dblY1)) + (dblY2 * (dblX1 - dblPointX)) - _
                (dblX1 * dblPointY) + (dblY1 * dblPointX)) > 0 Then lngWindCounter = lngWindCounter + 1
          If CalcCheckClockwiseNumbers2(dblX1, dblY1, dblX2, dblY2, dblPointX, dblPointY) > 0 Then lngWindCounter = lngWindCounter + 1
        End If
      Else  ' AUTOMATICALLY WE KNOW dblY1 > dblPointY; no test needed
        If dblY2 <= dblPointY Then  ' If a downward crossing
'           If ((dblX2 * (dblPointY - dblY1)) + (dblY2 * (dblX1 - dblPointX)) - _
                (dblX1 * dblPointY) + (dblY1 * dblPointX)) < 0 Then lngWindCounter = lngWindCounter - 1
          If CalcCheckClockwiseNumbers2(dblX1, dblY1, dblX2, dblY2, dblPointX, dblPointY) < 0 Then lngWindCounter = lngWindCounter - 1
        End If
      End If
      dblX1 = dblX2
      dblY1 = dblY2
    Next lngPointIndex
  Next lngRingIndex
  
  PointInPoly_Winding = lngWindCounter <> 0
  
End Function

Public Function PointInPoly_Crossing(dblPointX As Double, dblPointY As Double, varPolygon() As Variant, _
    Optional lngCrossCounter As Long) As Boolean

  ' ASSUMES POLYGON IS IN THE FORM OF A VARIANT ARRAY, WHERE EACH OBJECT IN THE ARRAY IS A POLYGON RING.
  ' EACH RING IS IN THE FORM OF A DOUBLE-ARRAY OF EACH VERTEX IN THE RING.
  ' VERTEX (0) = VERTEX (uBound(RingArray))

  ' adapted from http://geomalgorithms.com/a03-_inclusion.html
  '// cn_PnPoly(): crossing number test for a point in a polygon
  '//      Input:   P = a point,
  '//               V[] = vertex points of a polygon V[n+1] with V[n]=V[0]
  '//      Return:  0 = outside, 1 = inside
  '// This code is patterned after [Franklin, 2000]
  'int
  'cn_PnPoly( Point P, Point* V, int n )'{
  '    int    cn = 0;    // the  crossing number counter
  '
  '    // loop through all edges of the polygon
  '    for (int i=0; i<n; i++) {    // edge from V[i]  to V[i+1]
  '       if (((V[i].y <= P.y) && (V[i+1].y > P.y))     // an upward crossing
  '        || ((V[i].y > P.y) && (V[i+1].y <=  P.y))) { // a downward crossing
  '            // compute  the actual edge-ray intersect x-coordinate
  '            float vt = (float)(P.y  - V[i].y) / (V[i+1].y - V[i].y);
  '            if (P.x <  V[i].x + vt * (V[i+1].x - V[i].x)) // P.x < intersect
  '                 ++cn;   // a valid crossing of y=P.y right of P.x
  '        }
  '    }
  '    return (cn&1);    // 0 if even (out), and 1 if  odd (in)
  '
  '}
  '//===================================================================
  
  Dim lngRingIndex As Long
  Dim dblRing() As Double
  Dim lngPointIndex As Long
  
  ' TEST IF THIS IS FASTER
  ' YES; ROUGHLY 20% FASTER IF I DECLARE THESE VARIABLES RATHER THAN REFERENCE THE ARRAY DIRECLTY
  Dim dblX1 As Double
  Dim dblY1 As Double
  Dim dblX2 As Double
  Dim dblY2 As Double
  Dim dblSegProportion As Double
    
  lngCrossCounter = 0
  
  For lngRingIndex = 0 To UBound(varPolygon)
    dblRing = varPolygon(lngRingIndex)
    dblX1 = dblRing(0, 0)
    dblY1 = dblRing(1, 0)
    
    
    For lngPointIndex = 0 To UBound(dblRing, 2) - 1
      ' This "if-then" excludes all cases where the edge segment can't possibly intersect horizontal line, and
      ' also cases where edge segment itself is horizontal
      dblX2 = dblRing(0, lngPointIndex + 1)
      dblY2 = dblRing(1, lngPointIndex + 1)
'      Call MakeGraphicSegment(dblX1, dblX2, dblY1, dblY2)
      If (dblY1 <= dblPointY And dblY2 > dblPointY) Or (dblY1 > dblPointY And dblY2 <= dblPointY) Then
        dblSegProportion = (dblPointY - dblY1) / (dblY2 - dblY1)
        If dblPointX < dblX1 + (dblSegProportion * (dblX2 - dblX1)) Then lngCrossCounter = lngCrossCounter + 1
      End If

      dblX1 = dblX2
      dblY1 = dblY2
    Next lngPointIndex
    
  Next lngRingIndex
  PointInPoly_Crossing = lngCrossCounter Mod 2 = 1
  
End Function

Public Sub MakeGraphicSegment(dblX1 As Double, dblX2 As Double, dblY1 As Double, dblY2 As Double)
  
  Dim pSpRef As ISpatialReference
  Set pSpRef = MyGeneralOperations.CreateGeneralProjectedSpatialReference(3857)
  
  Dim pPoint As IPoint
  Dim pPolyline As IPolyline
  Dim pPtColl As IPointCollection
  
  Set pPolyline = New Polyline
  Set pPtColl = pPolyline
  Set pPoint = New Point
  pPoint.PutCoords dblX1, dblY1
  pPtColl.AddPoint pPoint
  Set pPoint = New Point
  pPoint.PutCoords dblX2, dblY2
  pPtColl.AddPoint pPoint
  Set pPolyline.SpatialReference = pSpRef
  
  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  
  MyGeneralOperations.DeleteGraphicsByName pMxDoc, "Delete_Me_Line"
  
  Dim pLineSym As ISimpleLineSymbol
  Set pLineSym = New SimpleLineSymbol
  pLineSym.Width = 2
  Dim pColor As IRgbColor
  Set pColor = New RgbColor
  pColor.RGB = RGB(255, 0, 255)
  pLineSym.Color = pColor
  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolyline, "Delete_Me_Line", pLineSym
  
ClearMemory:
  Set pSpRef = Nothing
  Set pPoint = Nothing
  Set pPolyline = Nothing
  Set pPtColl = Nothing
  Set pMxDoc = Nothing


  
End Sub

Public Function CalcCheckClockwiseNumbers2(dblPX As Double, dblPY As Double, dblQX As Double, _
    dblQY As Double, dblRX As Double, dblRY As Double, Optional dblDistance As Double) As Double
  ' CalcCheckClockwiseNumbers2
  ' Jenness Enterprises <www.jennessent.com)>
  ' Given 3 consecutive points, this scripts calculates whether the third point (R) lies to the right
  ' (clockwise) or to the left (counter-clockwise) of the line connecting the first point (P) to
  ' the second point (Q).
  ' RETURNS > 0 IF COUNTERCLOCKWISE
  '         = 0 IF ON LINE
  '         < 0 IF CLOCKWISE
 
  CalcCheckClockwiseNumbers2 = (dblQX * (dblRY - dblPY)) + (dblQY * (dblPX - dblRX)) - (dblPX * dblRY) _
        + (dblPY * dblRX)
 
End Function

Public Function FillPolygonWithArrayOfPointsForRaster(pPolygon As IPolygon, pRaster As IRaster, _
    dblCellSizeX As Double, dblCellSizeY As Double, dblRastMinX As Double, dblRastMinY As Double, dblRastMaxY As Double, _
    varPolygon() As Variant, lngPointCounter As Long, booArray() As Boolean, pOrigin As IPnt, _
    pPB As IPixelBlock, Optional booReturnCoords As Boolean = False, Optional varCoords As Variant, _
    Optional booUseHull As Boolean = False, Optional varHullPolyArray As Variant, _
    Optional pMxDoc As IMxDocument, Optional dblArrayEnvMinX As Double, _
    Optional dblArrayEnvMinY As Double) As Boolean

  Dim dblPoints() As Double
  Dim pEnv As IEnvelope
  Dim lngCounter As Long
  Dim booInsideWind As Boolean
  Dim dblMinX As Double
  Dim dblMinY As Double
  Dim dblMaxX As Double
  Dim dblMaxY As Double
  Dim dblIncX As Double
  Dim dblIncY As Double
  Dim dblShiftX As Double
  Dim dblShiftY As Double
  Dim dblShiftMaxX As Double
  Dim dblShiftMaxY As Double
  Dim dblWidth As Double
  Dim dblHeight As Double
  Dim varHullPoly() As Variant
    
  Set pEnv = pPolygon.Envelope
  pEnv.QueryCoords dblMinX, dblMinY, dblMaxX, dblMaxY
  dblWidth = pEnv.Width
  dblHeight = pEnv.Height
  
  dblShiftX = MyGeometricOperations.ModDouble((dblMinX - dblRastMinX), dblCellSizeX)
  dblShiftY = MyGeometricOperations.ModDouble((dblMinY - dblRastMinY), dblCellSizeY)
  dblShiftMaxX = MyGeometricOperations.ModDouble((dblMaxX - dblRastMinX), dblCellSizeX)
  dblShiftMaxY = MyGeometricOperations.ModDouble((dblMaxY - dblRastMinY), dblCellSizeY)
  
  ' BELOW GIVES COORDINATES OF LOWER LEFT CORNERS OF CELLS
  dblArrayEnvMinX = (dblMinX - dblShiftX)
  dblArrayEnvMinY = (dblMinY - dblShiftY)
  
  ' BELOW GIVES COORDINATES OF CELL CENTERS
  dblMinX = (dblMinX - dblShiftX) + (dblCellSizeX / 2)
  dblMinY = (dblMinY - dblShiftY) + (dblCellSizeY / 2)
  
  Dim dblX As Double
  Dim dblY As Double
  Dim lngX As Long
  Dim lngY As Long
  dblX = pEnv.XMin - dblRastMinX
  dblY = dblRastMaxY - pEnv.YMax
        
  'Find ncells from left-top
  'Determine cell index by taking the int. It will also take care of the fact that the index is 0 based
  lngX = Int(dblX / dblCellSizeX)
  lngY = Int(dblY / dblCellSizeY)
  Set pOrigin = New Pnt
  pOrigin.SetCoords lngX, lngY
'  iX = Int(nX)
'  iY = Int(nY)
  
  lngCounter = -1
  lngPointCounter = -1
  
  Dim lngRowMaxIndex As Long
  Dim lngColMaxIndex As Long
  lngRowMaxIndex = Int(pEnv.Width / dblCellSizeX)
  lngColMaxIndex = Int(pEnv.Height / dblCellSizeY)
  
  If dblShiftX > dblShiftMaxX Then lngRowMaxIndex = lngRowMaxIndex + 1
  If dblShiftY > dblShiftMaxY Then lngColMaxIndex = lngColMaxIndex + 1
  
'  Create a pnt to hold the 2x2 PixelBlock size
  Dim pPnt As IPnt
  Set pPnt = New Pnt
  pPnt.SetCoords lngRowMaxIndex + 1, lngColMaxIndex + 1
  Set pPB = pRaster.CreatePixelBlock(pPnt)
  
  Dim lngRowCounter As Long
  Dim lngColCounter As Long
  
  ' BOOLEAN ARRAY WILL HAVE ORIGIN AT LOWER LEFT CORNER.
  ' INDEX 0 = X-COORDINATE
  ' INDEX 1 = Y-COORDINATE
  ' ORIGONALLY THIS ARRAY WAS DIMMED TO lngRowMaxIndex, lngColMaxIndex
  ' RAN INTO INVALID INDEX ERRORS, POSSIBLY DUE TO ROUNDING ISSUES
  ' RESET IT TO ADD 1 TO EACH; MOST TIMES THIS WILL JUST RESULT IN THE CELLS BEING DECLARIED FALSE
  ReDim booArray(lngRowMaxIndex + 1, lngColMaxIndex + 1)
  
'  ' FOR DEBUGGING
'  Dim pPoint As IPoint
'  Dim pMxDoc As MxDocument
'  Set pMxDoc = ThisDocument
'
'  Dim pMarker1 As ISimpleMarkerSymbol
'  Dim pMarker2 As ISimpleMarkerSymbol
'  Dim pColor1 As IRgbColor
'  Dim pColor2 As IRgbColor
'  Set pColor1 = New RgbColor
'  Set pColor2 = New RgbColor
'
'  pColor1.RGB = RGB(255, 0, 0)
'  pColor2.RGB = RGB(0, 0, 255)
'
'  Set pMarker1 = New SimpleMarkerSymbol
'  Set pMarker2 = New SimpleMarkerSymbol
'
'  pMarker1.Style = esriSMSCircle
'  pMarker1.size = 8
'  pMarker1.Color = pColor1
'  pMarker2.Style = esriSMSCircle
'  pMarker2.size = 8
'  pMarker2.Color = pColor2
'  '==============================================
  
  Dim booInHull As Boolean
  
  lngRowCounter = -1
  For dblIncX = dblMinX To (dblMinX + dblShiftX + dblWidth) Step dblCellSizeX ' <-- THIS GOES RIGHT TO LEFT IN SPACE
    lngRowCounter = lngRowCounter + 1
    lngColCounter = lngColMaxIndex + 1
    For dblIncY = dblMinY To (dblMinY + dblShiftY + dblHeight) Step dblCellSizeY ' <-- THIS GOES BOTTOM TO TOP IN SPACE
      lngCounter = lngCounter + 1
      lngColCounter = lngColCounter - 1
      
      If booUseHull Then
        varHullPoly = varHullPolyArray(0)
        booInHull = PointInPoly_Winding(dblIncX, dblIncY, varHullPoly)
      Else
        booInHull = True
      End If
      If booInHull Then
        booInsideWind = PointInPoly_Winding(dblIncX, dblIncY, varPolygon)
        If booInsideWind Then
          lngPointCounter = lngPointCounter + 1
          If booReturnCoords Then
            ReDim Preserve dblPoints(1, lngPointCounter)
            dblPoints(0, lngPointCounter) = dblIncX
            dblPoints(1, lngPointCounter) = dblIncY
          End If
        End If
      Else
        booInsideWind = False
      End If
'       FOR DEBUGGING
'      Debug.Print "Inside = " & CStr(booInsideWind) & ", Cross Count = " & CStr(lngWindCount)
'      Set pPoint = New Point
'      Set pPoint.SpatialReference = pPolygon.SpatialReference
'      pPoint.PutCoords dblIncX, dblIncY
'      If booInsideWind Then
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint, "Inside", pMarker1
'      Else
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint, "Outside", pMarker2
'      End If
      
      booArray(lngRowCounter, lngColCounter) = booInsideWind
      
    Next dblIncY
  Next dblIncX
  
  If booReturnCoords Then
    ReDim varCoords(0)
    varCoords(0) = dblPoints
  End If
  
  FillPolygonWithArrayOfPointsForRaster = True
     
  
  
'  ' FOR DEBUGGING
'  Dim strReport As String
'  For lngColCounter = 0 To UBound(booArray, 2)
'    For lngRowCounter = 0 To UBound(booArray, 1)
'      strReport = strReport & IIf(booArray(lngRowCounter, lngColCounter), "X", ".")
'    Next lngRowCounter
'    strReport = strReport & vbCrLf
'  Next lngColCounter
'
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.Clear
'  pDataObj.SetText strReport
'  pDataObj.PutInClipboard
'  Set pDataObj = Nothing
  
ClearMemory:
  Erase dblPoints
  Set pEnv = Nothing

End Function
Public Function FillPolygonWithArrayOfIntCellsForRaster(pPolygon As IPolygon, pRaster As IRaster, _
    dblCellSizeX As Double, dblCellSizeY As Double, dblRastMinX As Double, dblRastMinY As Double, dblRastMaxY As Double, _
    varPolygon() As Variant, lngPointCounter As Long, booArray() As Boolean, pOrigin As IPnt, _
    pPB As IPixelBlock, varEnvelopes() As Variant, _
    Optional booReturnCoords As Boolean = False, Optional varCoords As Variant, _
    Optional booUseHull As Boolean = False, Optional varHullPolyArray As Variant, _
    Optional pMxDoc As IMxDocument) As Boolean
  
  ' This function similar to FillPolygonWithArrayOfPointsForRaster except that it fills
  ' array of envelope coordinates, which is used by the GridFunctions.CellValues2_Fast_byArray_VectorAdjust

  Dim dblPoints() As Double
  Dim pEnv As IEnvelope
  Dim lngCounter As Long
  Dim booInsideWind As Boolean
  Dim dblMinX As Double
  Dim dblMinY As Double
  Dim dblMaxX As Double
  Dim dblMaxY As Double
  Dim dblIncX As Double
  Dim dblIncY As Double
  Dim dblShiftX As Double
  Dim dblShiftY As Double
  Dim dblShiftMaxX As Double
  Dim dblShiftMaxY As Double
  Dim dblWidth As Double
  Dim dblHeight As Double
  Dim varHullPoly() As Variant
    
  Set pEnv = pPolygon.Envelope
  pEnv.QueryCoords dblMinX, dblMinY, dblMaxX, dblMaxY
  dblWidth = pEnv.Width
  dblHeight = pEnv.Height
  
  dblShiftX = MyGeometricOperations.ModDouble((dblMinX - dblRastMinX), dblCellSizeX)
  dblShiftY = MyGeometricOperations.ModDouble((dblMinY - dblRastMinY), dblCellSizeY)
  dblShiftMaxX = MyGeometricOperations.ModDouble((dblMaxX - dblRastMinX), dblCellSizeX)
  dblShiftMaxY = MyGeometricOperations.ModDouble((dblMaxY - dblRastMinY), dblCellSizeY)
  
  ' BELOW GIVES COORDINATES OF CELL CENTERS
  dblMinX = (dblMinX - dblShiftX) + (dblCellSizeX / 2)
  dblMinY = (dblMinY - dblShiftY) + (dblCellSizeY / 2)
  
  Dim dblX As Double
  Dim dblY As Double
  Dim lngX As Long
  Dim lngY As Long
  dblX = pEnv.XMin - dblRastMinX
  dblY = dblRastMaxY - pEnv.YMax
        
  'Find ncells from left-top
  'Determine cell index by taking the int. It will also take care of the fact that the index is 0 based
  lngX = Int(dblX / dblCellSizeX)
  lngY = Int(dblY / dblCellSizeY)
  Set pOrigin = New Pnt
  pOrigin.SetCoords lngX, lngY
'  iX = Int(nX)
'  iY = Int(nY)
  
  lngCounter = -1
  lngPointCounter = -1
  
  Dim lngRowMaxIndex As Long
  Dim lngColMaxIndex As Long
  lngRowMaxIndex = Int(pEnv.Width / dblCellSizeX)
  lngColMaxIndex = Int(pEnv.Height / dblCellSizeY)
  
  If dblShiftX > dblShiftMaxX Then lngRowMaxIndex = lngRowMaxIndex + 1
  If dblShiftY > dblShiftMaxY Then lngColMaxIndex = lngColMaxIndex + 1
  
'  Create a pnt to hold the 2x2 PixelBlock size
  Dim pPnt As IPnt
  Set pPnt = New Pnt
  pPnt.SetCoords lngRowMaxIndex + 1, lngColMaxIndex + 1
  Set pPB = pRaster.CreatePixelBlock(pPnt)
  
  Dim lngRowCounter As Long
  Dim lngColCounter As Long
  
  ' BOOLEAN ARRAY WILL HAVE ORIGIN AT LOWER LEFT CORNER.
  ' INDEX 0 = X-COORDINATE
  ' INDEX 1 = Y-COORDINATE
  ' ORIGONALLY THIS ARRAY WAS DIMMED TO lngRowMaxIndex, lngColMaxIndex
  ' RAN INTO INVALID INDEX ERRORS, POSSIBLY DUE TO ROUNDING ISSUES
  ' RESET IT TO ADD 1 TO EACH; MOST TIMES THIS WILL JUST RESULT IN THE CELLS BEING DECLARIED FALSE
  ReDim booArray(lngRowMaxIndex + 1, lngColMaxIndex + 1)
  ReDim varEnvelopes(lngRowMaxIndex + 1, lngColMaxIndex + 1)
  
'  ' FOR DEBUGGING
'  Dim pPoint As IPoint
'
'  Dim pMarker1 As ISimpleMarkerSymbol
'  Dim pMarker2 As ISimpleMarkerSymbol
'  Dim pColor1 As IRgbColor
'  Dim pColor2 As IRgbColor
'  Set pColor1 = New RgbColor
'  Set pColor2 = New RgbColor
'
'  pColor1.RGB = RGB(255, 0, 0)
'  pColor2.RGB = RGB(0, 0, 255)
'
'  Set pMarker1 = New SimpleMarkerSymbol
'  Set pMarker2 = New SimpleMarkerSymbol
'
'  pMarker1.Style = esriSMSCircle
'  pMarker1.size = 8
'  pMarker1.Color = pColor1
'  pMarker2.Style = esriSMSCircle
'  pMarker2.size = 8
'  pMarker2.Color = pColor2
'  '==============================================
  
  Dim booInHull As Boolean
  Dim pRelOp As IRelationalOperator
  Dim pTestEnv As IEnvelope
  
  Set pRelOp = pPolygon
  
  ' FOR INTERSECTING CELLS
  
  lngRowCounter = -1
  For dblIncX = dblMinX To (dblMinX + dblShiftX + dblWidth) Step dblCellSizeX ' <-- THIS GOES RIGHT TO LEFT IN SPACE
    lngRowCounter = lngRowCounter + 1
    lngColCounter = lngColMaxIndex + 1
'    Debug.Print "Row " & CStr(lngRowCounter) & " of " & CStr(lngRowMaxIndex)
    For dblIncY = dblMinY To (dblMinY + dblShiftY + dblHeight) Step dblCellSizeY ' <-- THIS GOES BOTTOM TO TOP IN SPACE
      lngCounter = lngCounter + 1
      lngColCounter = lngColCounter - 1
    
      Set pTestEnv = New Envelope
      Set pTestEnv.SpatialReference = pPolygon.SpatialReference
      pTestEnv.PutCoords dblIncX - (dblCellSizeX / 2), dblIncY - (dblCellSizeY / 2), _
          dblIncX + (dblCellSizeX / 2), dblIncY + (dblCellSizeY / 2)
      
      If booUseHull Then
        varHullPoly = varHullPolyArray(0)
        booInHull = PointInPoly_Winding(dblIncX, dblIncY, varHullPoly)
      Else
        booInHull = True
      End If
      If booInHull Then
        booInsideWind = PointInPoly_Winding(dblIncX, dblIncY, varPolygon)
        If booInsideWind Then
          lngPointCounter = lngPointCounter + 1
          If booReturnCoords Then
            ReDim Preserve dblPoints(1, lngPointCounter)
            dblPoints(0, lngPointCounter) = dblIncX
            dblPoints(1, lngPointCounter) = dblIncY
          End If
        End If
      Else
        booInsideWind = False
      End If
      
      ' ALTERNATIVE TEST FOR INCLUDING INTERSECTING EDGE CELLS
      If Not booInsideWind Then
        booInsideWind = Not pRelOp.Disjoint(pTestEnv)
        If booInsideWind Then
          lngPointCounter = lngPointCounter + 1
          If booReturnCoords Then
            ReDim Preserve dblPoints(1, lngPointCounter)
            dblPoints(0, lngPointCounter) = dblIncX
            dblPoints(1, lngPointCounter) = dblIncY
          End If
        End If
      End If
      
''       FOR DEBUGGING
'      Debug.Print "Inside = " & CStr(booInsideWind) & ", Cross Count = " & CStr(lngPointCounter)
'      Set pPoint = New Point
'      Set pPoint.SpatialReference = pPolygon.SpatialReference
'      pPoint.PutCoords dblIncX, dblIncY
'      If booInsideWind Then
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint, "Inside", pMarker1
'      Else
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint, "Outside", pMarker2
'      End If

      booArray(lngRowCounter, lngColCounter) = booInsideWind
      Set varEnvelopes(lngRowCounter, lngColCounter) = pTestEnv
      
    Next dblIncY
  Next dblIncX
  
  If booReturnCoords Then
    ReDim varCoords(0)
    varCoords(0) = dblPoints
  End If
  
  FillPolygonWithArrayOfIntCellsForRaster = True
     
  
  
'  ' FOR DEBUGGING
'  Dim strReport As String
'  For lngColCounter = 0 To UBound(booArray, 2)
'    For lngRowCounter = 0 To UBound(booArray, 1)
'      strReport = strReport & IIf(booArray(lngRowCounter, lngColCounter), "X", ".")
'    Next lngRowCounter
'    strReport = strReport & vbCrLf
'  Next lngColCounter
'
'  Dim pDataObj As New MSForms.DataObject
'  pDataObj.Clear
'  pDataObj.SetText strReport
'  pDataObj.PutInClipboard
'  Set pDataObj = Nothing
  
ClearMemory:
  Erase dblPoints
  Set pEnv = Nothing

End Function


Public Function SubdividePolygon(pPolygon As IPolygon, lngMaxVertexCount As Long, _
      Optional pMxDoc As IMxDocument, Optional dblCellSizeX As Double = -999, _
      Optional dblCellSizeY As Double = -999) As Variant()

  Dim pPtColl As IPointCollection
  Dim lngPtColl As Long
  Dim pTopoOp As ITopologicalOperator4
  Dim pGeomBag As IGeometryBag
  Dim varReturn() As Variant
  Dim lngReturnCounter As Long
'  Dim varCurrentPolys() As Variant
  Dim pSplitLine As IPolyline
  Dim booNeedsSplit As Boolean
  Dim lngIndex As Long
  Dim pTestPoly As IPolygon
  Dim pEnv As IEnvelope
'  Dim pSplitArray As esriSystem.IArray
  Dim lngSplitCounter As Long
  Dim pEnumGeom As IEnumGeometry
  Dim pSubPoly As IPolygon
  Dim pSubPoly2 As IPolygon
  Dim booSucceeded As Boolean
  Dim strFailureReason As String
  Dim varGoodArray() As Variant
  Dim varBadArray() As Variant
  Dim varTempArray() As Variant
  Dim lngGoodCounter As Long
  Dim lngBadCounter As Long
  
  Dim booCheckPolygonCount As Boolean
  booCheckPolygonCount = dblCellSizeX > 0 And dblCellSizeY > 0
  
  lngGoodCounter = -1
  lngBadCounter = -1
  Dim dblWidth As Double
  Dim dblHeight As Double
  
  If pPolygon.IsEmpty Then
    ReDim varReturn(0)
    varReturn(0) = pPolygon
    SubdividePolygon = varReturn
    GoTo ClearMemory
  Else
    varTempArray = SplitMultipartFeatureIntoVarArray(pPolygon, booSucceeded, strFailureReason)
    
    For lngIndex = 0 To UBound(varTempArray)
      Set pPolygon = varTempArray(lngIndex)
      Set pEnv = pPolygon.Envelope
      dblWidth = pEnv.Width
      dblHeight = pEnv.Height
      Set pPtColl = pPolygon
      If pPtColl.PointCount >= lngMaxVertexCount Then
        lngBadCounter = lngBadCounter + 1
        ReDim Preserve varBadArray(lngBadCounter)
        Set varBadArray(lngBadCounter) = pPolygon
      Else
        If booCheckPolygonCount Then
          If (dblWidth / dblCellSizeX) * (dblHeight / dblCellSizeY) > 750000 Then
            lngBadCounter = lngBadCounter + 1
            ReDim Preserve varBadArray(lngBadCounter)
            Set varBadArray(lngBadCounter) = pPolygon
          Else
            lngGoodCounter = lngGoodCounter + 1
            ReDim Preserve varGoodArray(lngGoodCounter)
            Set varGoodArray(lngGoodCounter) = pPolygon
          End If
        Else
          lngGoodCounter = lngGoodCounter + 1
          ReDim Preserve varGoodArray(lngGoodCounter)
          Set varGoodArray(lngGoodCounter) = pPolygon
        End If
      End If
    Next lngIndex
    If lngBadCounter = -1 Then
      SubdividePolygon = varTempArray
      GoTo ClearMemory
    End If
  End If
  
'  ReDim varCurrentPolys(0)
'  Set varCurrentPolys(0) = pPolygon
  
'  varCurrentPolys = varBadArray
  Dim varRunningArray() As Variant
  Dim lngRunningCounter As Long
  lngReturnCounter = -1
  Dim varSplitArray() As Variant
  
  
  booNeedsSplit = True
  Do Until booNeedsSplit = False
    booNeedsSplit = False
    lngRunningCounter = -1
'    Debug.Print
    
    For lngIndex = 0 To UBound(varBadArray)
      Set pTestPoly = varBadArray(lngIndex)
      ' WE KNOW THIS POLYGON HAS TOO MANY VERTICES BECAUSE IT IS FROM THE BAD ARRAY
      Set pEnv = pTestPoly.Envelope
      Set pSplitLine = ReturnEnvelopeMidLine(pEnv)
      
      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pSplitLine, "Delete_Me"
      
      Set pTopoOp = pTestPoly
      Set pGeomBag = pTopoOp.Cut2(pSplitLine)
      Set pEnumGeom = pGeomBag
      pEnumGeom.Reset
      Set pSubPoly = pEnumGeom.Next
      Do Until pSubPoly Is Nothing
        
        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pSubPoly, "Delete_Me"
        
        varSplitArray = MyGeometricOperations.SplitMultipartFeatureIntoVarArray(pSubPoly, True, "")
        For lngSplitCounter = 0 To UBound(varSplitArray) ' pSplitArray.Count - 1
          Set pSubPoly2 = varSplitArray(lngSplitCounter) ' pSplitArray.Element(lngSplitCounter)
          pSubPoly2.SimplifyPreserveFromTo
          
          MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pSubPoly2, "Delete_Me"
        
          Set pEnv = pSubPoly2.Envelope
          dblWidth = pEnv.Width
          dblHeight = pEnv.Height
          Set pPtColl = pSubPoly2
          If pPtColl.PointCount <= lngMaxVertexCount Then
            If booCheckPolygonCount Then
              If (dblWidth / dblCellSizeX) * (dblHeight / dblCellSizeY) > 750000 Then
                lngRunningCounter = lngRunningCounter + 1
                ReDim Preserve varRunningArray(lngRunningCounter)
                Set varRunningArray(lngRunningCounter) = pSubPoly2
              Else
                lngGoodCounter = lngGoodCounter + 1
                ReDim Preserve varGoodArray(lngGoodCounter)
                Set varGoodArray(lngGoodCounter) = pSubPoly2
              End If
            Else
              lngGoodCounter = lngGoodCounter + 1
              ReDim Preserve varGoodArray(lngGoodCounter)
              Set varGoodArray(lngGoodCounter) = pSubPoly2
            End If
            
          Else
            booNeedsSplit = True
            lngRunningCounter = lngRunningCounter + 1
            ReDim Preserve varRunningArray(lngRunningCounter)
            Set varRunningArray(lngRunningCounter) = pSubPoly2
'            Debug.Print "  --> [" & CStr(lngRunningCounter) & "] " & Format(pPtColl.PointCount, "#,##0") & " vertices..."
          End If
        Next lngSplitCounter
        
        Set pSubPoly = pEnumGeom.Next
      Loop
    Next lngIndex
    
    varBadArray = varRunningArray
  Loop
  
  SubdividePolygon = varGoodArray
  
ClearMemory:
  Set pPtColl = Nothing
  Set pTopoOp = Nothing
  Set pGeomBag = Nothing
  Erase varReturn
  Set pSplitLine = Nothing
  Set pTestPoly = Nothing
  Set pEnv = Nothing
  Set pEnumGeom = Nothing
  Set pSubPoly = Nothing
  Set pSubPoly2 = Nothing
  Erase varGoodArray
  Erase varBadArray
  Erase varTempArray
  Erase varRunningArray
  Erase varSplitArray



End Function
Public Function SubdividePolygon2(pPolygon As IPolygon, lngMaxVertexCount As Long, _
      Optional pMxDoc As IMxDocument, Optional dblCellSizeX As Double = -999, _
      Optional dblCellSizeY As Double = -999) As Variant()

  Dim pPtColl As IPointCollection
  Dim lngPtColl As Long
  Dim pTopoOp As ITopologicalOperator4
'  Dim pGeomBag As IGeometryBag
  Dim varReturn() As Variant
  Dim lngReturnCounter As Long
'  Dim varCurrentPolys() As Variant
  Dim pSplitLine As IPolyline
  Dim booNeedsSplit As Boolean
  Dim lngIndex As Long
  Dim pTestPoly As IPolygon
  Dim pEnv As IEnvelope
'  Dim pSplitArray As esriSystem.IArray
  Dim lngSplitCounter As Long
'  Dim pEnumGeom As IEnumGeometry
  Dim pSubPoly As IPolygon
  Dim pSubPoly2 As IPolygon
  Dim booSucceeded As Boolean
  Dim strFailureReason As String
  Dim varGoodArray() As Variant
  Dim varBadArray() As Variant
  Dim varTempArray() As Variant
  Dim lngGoodCounter As Long
  Dim lngBadCounter As Long
  
  Dim pLeftPoly As IPolygon
  Dim pRightPoly As IPolygon
  
  Dim booCheckPolygonCount As Boolean
  booCheckPolygonCount = dblCellSizeX > 0 And dblCellSizeY > 0
  
  lngGoodCounter = -1
  lngBadCounter = -1
  Dim dblWidth As Double
  Dim dblHeight As Double
  
  If pPolygon.IsEmpty Then
    ReDim varReturn(0)
    varReturn(0) = pPolygon
    SubdividePolygon2 = varReturn
    GoTo ClearMemory
  Else
    varTempArray = SplitMultipartFeatureIntoVarArray(pPolygon, booSucceeded, strFailureReason)
    
    For lngIndex = 0 To UBound(varTempArray)
      Set pPolygon = varTempArray(lngIndex)
      Set pEnv = pPolygon.Envelope
      dblWidth = pEnv.Width
      dblHeight = pEnv.Height
      Set pPtColl = pPolygon
      If pPtColl.PointCount >= lngMaxVertexCount Then
        lngBadCounter = lngBadCounter + 1
        ReDim Preserve varBadArray(lngBadCounter)
        Set varBadArray(lngBadCounter) = pPolygon
      Else
        If booCheckPolygonCount Then
          If (dblWidth / dblCellSizeX) * (dblHeight / dblCellSizeY) > 750000 Then
            lngBadCounter = lngBadCounter + 1
            ReDim Preserve varBadArray(lngBadCounter)
            Set varBadArray(lngBadCounter) = pPolygon
          Else
            lngGoodCounter = lngGoodCounter + 1
            ReDim Preserve varGoodArray(lngGoodCounter)
            Set varGoodArray(lngGoodCounter) = pPolygon
          End If
        Else
          lngGoodCounter = lngGoodCounter + 1
          ReDim Preserve varGoodArray(lngGoodCounter)
          Set varGoodArray(lngGoodCounter) = pPolygon
        End If
      End If
    Next lngIndex
    If lngBadCounter = -1 Then
      SubdividePolygon2 = varTempArray
      GoTo ClearMemory
    End If
  End If
  
'  ReDim varCurrentPolys(0)
'  Set varCurrentPolys(0) = pPolygon
  
'  varCurrentPolys = varBadArray
  Dim varRunningArray() As Variant
  Dim lngRunningCounter As Long
  lngReturnCounter = -1
  Dim varSplitArray() As Variant
  Dim varPair() As Variant
  Dim lngPairIndex As Long
  
  booNeedsSplit = True
  Do Until booNeedsSplit = False
    booNeedsSplit = False
    lngRunningCounter = -1
'    Debug.Print
    
    For lngIndex = 0 To UBound(varBadArray)
      Set pTestPoly = varBadArray(lngIndex)
      ' WE KNOW THIS POLYGON HAS TOO MANY VERTICES BECAUSE IT IS FROM THE BAD ARRAY
      Set pEnv = pTestPoly.Envelope
      Set pSplitLine = ReturnEnvelopeMidLine(pEnv)
      
'      MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pSplitLine, "Delete_Me"
      
      Set pTopoOp = pTestPoly
      pTopoOp.Cut pSplitLine, pLeftPoly, pRightPoly
      ReDim varPair(1)
      Set varPair(0) = pLeftPoly
      Set varPair(1) = pRightPoly
'
'      Set pGeomBag = pTopoOp.Cut2(pSplitLine)
'      Set pEnumGeom = pGeomBag
'      pEnumGeom.Reset
'      Set pSubPoly = pEnumGeom.Next
'      Do Until pSubPoly Is Nothing
      
      For lngPairIndex = 0 To UBound(varPair)
        Set pSubPoly = varPair(lngPairIndex)
        
'        MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pSubPoly, "Delete_Me"
        
        varSplitArray = MyGeometricOperations.SplitMultipartFeatureIntoVarArray(pSubPoly, True, "")
        For lngSplitCounter = 0 To UBound(varSplitArray) ' pSplitArray.Count - 1
          Set pSubPoly2 = varSplitArray(lngSplitCounter) ' pSplitArray.Element(lngSplitCounter)
          pSubPoly2.SimplifyPreserveFromTo
          
'          MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pSubPoly2, "Delete_Me"
        
          Set pEnv = pSubPoly2.Envelope
          dblWidth = pEnv.Width
          dblHeight = pEnv.Height
          Set pPtColl = pSubPoly2
          If pPtColl.PointCount <= lngMaxVertexCount Then
            If booCheckPolygonCount Then
              If (dblWidth / dblCellSizeX) * (dblHeight / dblCellSizeY) > 750000 Then
                lngRunningCounter = lngRunningCounter + 1
                ReDim Preserve varRunningArray(lngRunningCounter)
                Set varRunningArray(lngRunningCounter) = pSubPoly2
              Else
                lngGoodCounter = lngGoodCounter + 1
                ReDim Preserve varGoodArray(lngGoodCounter)
                Set varGoodArray(lngGoodCounter) = pSubPoly2
              End If
            Else
              lngGoodCounter = lngGoodCounter + 1
              ReDim Preserve varGoodArray(lngGoodCounter)
              Set varGoodArray(lngGoodCounter) = pSubPoly2
            End If
            
          Else
            booNeedsSplit = True
            lngRunningCounter = lngRunningCounter + 1
            ReDim Preserve varRunningArray(lngRunningCounter)
            Set varRunningArray(lngRunningCounter) = pSubPoly2
'            Debug.Print "  --> [" & CStr(lngRunningCounter) & "] " & Format(pPtColl.PointCount, "#,##0") & " vertices..."
          End If
        Next lngSplitCounter
        
'        Set pSubPoly = pEnumGeom.Next
'      Loop
      Next lngPairIndex
    Next lngIndex
    
    varBadArray = varRunningArray
  Loop
  
  SubdividePolygon2 = varGoodArray
  
ClearMemory:
  Set pPtColl = Nothing
  Set pTopoOp = Nothing
  Erase varReturn
  Set pSplitLine = Nothing
  Set pTestPoly = Nothing
  Set pEnv = Nothing
  Set pSubPoly = Nothing
  Set pSubPoly2 = Nothing
  Erase varGoodArray
  Erase varBadArray
  Erase varTempArray
  Set pLeftPoly = Nothing
  Set pRightPoly = Nothing
  Erase varRunningArray
  Erase varSplitArray
  Erase varPair



End Function


Public Function ReturnEnvelopeMidLine(pEnv As IEnvelope) As IPolyline
  
'  Redo this to split at middle or mean X/Y value!  Was getting nearly 500 polygons when splitting by geographic midpoint, and probably
'  doing lots of extra splits.
  
  Dim pPoint As IPoint
  Dim pPolyline As IPolyline
  Dim pPtColl As IPointCollection
  Set pPtColl = New Polyline
  Set pPolyline = pPtColl
  Set pPolyline.SpatialReference = pEnv.SpatialReference
  Dim dblXMin As Double
  Dim dblYMin As Double
  Dim dblXMax As Double
  Dim dblYMax As Double
  
  pEnv.QueryCoords dblXMin, dblYMin, dblXMax, dblYMax
  
  If pEnv.Width > pEnv.Height Then   ' LANDSCAPE
    Set pPoint = New Point
    Set pPoint.SpatialReference = pEnv.SpatialReference
    pPoint.PutCoords dblXMin + (pEnv.Width / 2), dblYMin - (pEnv.Height / 20)
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    Set pPoint.SpatialReference = pEnv.SpatialReference
    pPoint.PutCoords dblXMin + (pEnv.Width / 2), dblYMax + (pEnv.Height / 20)
    pPtColl.AddPoint pPoint
  Else                               ' PORTRAIT
    Set pPoint = New Point
    Set pPoint.SpatialReference = pEnv.SpatialReference
    pPoint.PutCoords dblXMin - (pEnv.Width / 20), dblYMin + (pEnv.Height / 2)
    pPtColl.AddPoint pPoint
    Set pPoint = New Point
    Set pPoint.SpatialReference = pEnv.SpatialReference
    pPoint.PutCoords dblXMax + (pEnv.Width / 20), dblYMin + (pEnv.Height / 2)
    pPtColl.AddPoint pPoint
  End If
  
  Set ReturnEnvelopeMidLine = pPolyline
  
ClearMemory:
  Set pPoint = Nothing
  Set pPolyline = Nothing
  Set pPtColl = Nothing

End Function

Public Sub TestClipProportion()

  Dim pMxDoc As IMxDocument
  Set pMxDoc = ThisDocument
  Dim pFLayer As IFeatureLayer
  Dim pAOIPolygon As IPolygon
  Dim pAOIFClass As IFeatureClass
  Dim pFeature As IFeature
  Dim pAnalysisPolygon As IPolygon
  
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("GAUL Administrative Units", pMxDoc.FocusMap)
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("hydrobasins_world", pMxDoc.FocusMap)
  Set pAOIFClass = pFLayer.FeatureClass
'  Set pAOIPolygon = pAOIFClass.GetFeature(33800).ShapeCopy  ' FULL COVERAGE
  Set pAOIPolygon = pAOIFClass.GetFeature(32832).ShapeCopy  ' PARTIAL COVERAGE
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pAOIPolygon, "Delete_Me"
  
  Set pFLayer = MyGeneralOperations.ReturnLayerByName("GPW Raster Cells", pMxDoc.FocusMap)
'  Set pAnalysisPolygon = pFLayer.FeatureClass.GetFeature(217).ShapeCopy  ' FULL COVERAGE
  Set pAnalysisPolygon = pFLayer.FeatureClass.GetFeature(6056).ShapeCopy  ' PARTIAL COVERAGE
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pClipPolygon, "Delete_Me"
  
  Dim dblFullArea As Double
  Dim dblClipArea As Double
  Dim dblProportion As Double
  Dim dblPolygonSumArea As Double
  Dim pClipPoly As IPolygon
  Dim varReturn As Variant
  
  varReturn = ReturnClipAndProportion(pAOIPolygon, pAnalysisPolygon, pAOIFClass)
  Set pClipPoly = varReturn(0)
  dblFullArea = varReturn(1)
  dblPolygonSumArea = varReturn(2)
  dblClipArea = varReturn(3)
  dblProportion = varReturn(4)
  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pClipPoly, "Delete_Me"
  
  Debug.Print "---------------------------"
  Debug.Print "Full Area = " & Format(dblFullArea / 1000000, "0.000") & " Sq. Km."
  Debug.Print "Full Polygon Area = " & Format(dblPolygonSumArea / 1000000, "0.000") & " Sq. Km."
  Debug.Print "Clipped Area = " & Format(dblClipArea / 1000000, "0.000") & " Sq. Km."
  Debug.Print "Proportion = " & Format(dblProportion, "0.0%")
  Debug.Print "Done..."
  
ClearMemory:
  Set pMxDoc = Nothing
  Set pFLayer = Nothing
  Set pAOIPolygon = Nothing
  Set pFeature = Nothing
  Set pAnalysisPolygon = Nothing
  Set pClipPoly = Nothing
  varReturn = Null


End Sub

Public Function ReturnClipAndProportion(pAOIPolygon As IPolygon, pPolygonToClip As IPolygon, _
    pFClass As IFeatureClass, Optional pMxDoc As IMxDocument, Optional booPause As Boolean, _
    Optional booForceFullArea As Boolean, Optional booUseDifferentArea As Boolean, _
    Optional pTileFClass As IFeatureClass, _
    Optional pTransform As IGeoTransformation, Optional pTransformDir As esriTransformDirection) _
    As Variant
    
  ' RETURNS VARIANT ARRAY WITH 4 ITEMS:
  ' 0) CLIPPED POLYGON
  ' 1) FULL AREA OF pPolygonToClip, i.e. Rectangle
  ' 2) FULL AREA OF REGION WITHIN pPolygonToClip COVERED BY ANY POLYGON FROM FEATURE CLASS
  ' 3) AREA OF CLIPPED PORTION OF pPolygonToClip
  ' 4) PROPORTION, RANGING FROM 0 TO 1
  '
  '' TEST FUNCTION
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'  Dim pFLayer As IFeatureLayer
'  Dim pAOIPolygon As IPolygon
'  Dim pAOIFClass As IFeatureClass
'  Dim pFeature As IFeature
'  Dim pAnalysisPolygon As IPolygon
'
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("GAUL Administrative Units", pMxDoc.FocusMap)
'  Set pAOIFClass = pFLayer.FeatureClass
''  Set pAOIPolygon = pAOIFClass.GetFeature(33800).ShapeCopy  ' FULL COVERAGE
'  Set pAOIPolygon = pAOIFClass.GetFeature(33799).ShapeCopy  ' PARTIAL COVERAGE
''  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pAOIPolygon, "Delete_Me"
'
'  Set pFLayer = MyGeneralOperations.ReturnLayerByName("GPW Raster Cells", pMxDoc.FocusMap)
''  Set pAnalysisPolygon = pFLayer.FeatureClass.GetFeature(217).ShapeCopy  ' FULL COVERAGE
'  Set pAnalysisPolygon = pFLayer.FeatureClass.GetFeature(180).ShapeCopy  ' PARTIAL COVERAGE
''  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pClipPolygon, "Delete_Me"
'
'  Dim dblFullArea As Double
'  Dim dblClipArea As Double
'  Dim dblProportion As Double
'  Dim dblPolygonSumArea As Double
'  Dim pClipPoly As IPolygon
'  Dim varReturn As Variant
'
'  varReturn = ReturnClipAndProportion(pAOIPolygon, pAnalysisPolygon, pAOIFClass)
'  Set pClipPoly = varReturn(0)
'  dblFullArea = varReturn(1)
'  dblPolygonSumArea = varReturn(2)
'  dblClipArea = varReturn(3)
'  dblProportion = varReturn(4)
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pClipPoly, "Delete_Me"
'
'  Debug.Print "---------------------------"
'  Debug.Print "Full Area = " & Format(dblFullArea / 1000000, "0.000") & " Sq. Km."
'  Debug.Print "Full Polygon Area = " & Format(dblPolygonSumArea / 1000000, "0.000") & " Sq. Km."
'  Debug.Print "Clipped Area = " & Format(dblClipArea / 1000000, "0.000") & " Sq. Km."
'  Debug.Print "Proportion = " & Format(dblProportion, "0.0%")
'  Debug.Print "Done..."
'
'ClearMemory:
'  Set pMxDoc = Nothing
'  Set pFLayer = Nothing
'  Set pAOIPolygon = Nothing
'  Set pFeature = Nothing
'  Set pAnalysisPolygon = Nothing
'  Set pClipPoly = Nothing
'  varReturn = Null

  Dim pClone As IClone
  Dim pAnalysisPolygon As IPolygon
  Dim pGeom As IGeometry2
  Dim booShouldProject As Boolean
  Dim booSREqual As Boolean
  Dim dblSemiMajorAxis As Double
  Dim dblSemiMinorAxis As Double
  Dim pSpheroid As ISpheroid
  Dim pGeoSpRef As IGeographicCoordinateSystem
  
  If booPause Then
    DoEvents
  End If
  ' PROJECT ENVELOPE (pPolygonToClip) TO AOI SPATIAL REFERENCE
  Set pGeoSpRef = pAOIPolygon.SpatialReference  ' MAY OR MAY NOT BE GEOGRAPHIC SPATIAL REFERENCE
  
  booShouldProject = Not MyGeneralOperations.CompareSpatialReferences(pGeoSpRef, _
    pPolygonToClip.SpatialReference, booSREqual)
  
  Set pClone = pPolygonToClip
  Set pAnalysisPolygon = pClone.Clone
  Set pGeom = pAnalysisPolygon
  If booShouldProject Then
    If Not pTransform Is Nothing Then
      pGeom.ProjectEx pGeoSpRef, pTransformDir, pTransform, False, 0, 0
    Else
      pGeom.Project pGeoSpRef
    End If
  End If
  
  
  ' CALCULATE TOTAL AREA IN ENVELOPE COVERED BY POLYGONS FROM AOI FCLASS
  Dim pTopoOp As ITopologicalOperator
  Dim pRelOp As IRelationalOperator2
  Dim pFCursor As IFeatureCursor
  Dim pFeature As IFeature
  Dim pSubPoly As IPolygon
  Dim pSubClip As IPolygon4
  Dim pSpFilt As ISpatialFilter
  Dim dblPolygonSumArea As Double
  Dim dblFullArea As Double
  Dim dblClipArea As Double
  Dim pArea As IArea
  Dim booUseSpherical As Boolean
  Dim pSubPolyArray As esriSystem.IVariantArray
  Set pSubPolyArray = New esriSystem.varArray
  Dim pCombinePolygon As IPolygon
  Dim dblProportion As Double
  Dim pTopoOp4 As ITopologicalOperator4
  
  booUseSpherical = TypeOf pGeoSpRef Is IGeographicCoordinateSystem
  
  Set pSpFilt = New SpatialFilter
  pSpFilt.SpatialRel = esriSpatialRelIntersects
  Set pSpFilt.Geometry = pAnalysisPolygon
'  Set pFCursor = pTileFClass.Search(pSpFilt, False)
  Set pFCursor = pFClass.Search(pSpFilt, False)
  
  ' FOR WEIRD ARCOBJECTS BUG
  Dim pProjPoly As IPolygon
  Dim pTempPoly As IPolygon
  Dim booIsGeographic As Boolean
  Dim pTempGeomColl As IGeometryCollection
  Dim pProjSpRef As IProjectedCoordinateSystem
  Dim pGeoSpRef2 As IGeographicCoordinateSystem
  If TypeOf pAnalysisPolygon.SpatialReference Is IGeographicCoordinateSystem Then
    Set pGeoSpRef2 = pAnalysisPolygon.SpatialReference
    booIsGeographic = True
  Else
    Set pProjSpRef = pAnalysisPolygon.SpatialReference
    Set pGeoSpRef2 = pProjSpRef.GeographicCoordinateSystem
    booIsGeographic = False
  End If
  
  Dim pClipPoly As IPolygon
  Set pClone = pAnalysisPolygon
  Set pClipPoly = pClone.Clone
  Dim pClipEnv As IEnvelope
  Set pClipEnv = pClipPoly.Envelope
  If booUseDifferentArea Then
    pClipEnv.Expand 3.5, 2.5, True
  Else
    pClipEnv.Expand 2, 2, True
  End If
  Set pClipPoly = EnvelopeToPolygon(pClipEnv)
  
  Set pTopoOp = pClipPoly
  
  If booForceFullArea Then
    Set pClone = pAnalysisPolygon
    Set pCombinePolygon = pClone.Clone
  Else
    Set pFeature = pFCursor.NextFeature
    Do Until pFeature Is Nothing
      Set pSubPoly = pFeature.ShapeCopy
      Set pRelOp = pSubPoly
      
        If pRelOp.Contains(pAnalysisPolygon) Then
          Set pSubClip = pAnalysisPolygon
        Else
          Set pSubClip = pTopoOp.Intersect(pSubPoly, pSubPoly.Dimension)
        End If
      
      ' FOR WEIRD ARCOBJECTS BUG
      If booUseDifferentArea Then
        Set pTempGeomColl = New GeometryBag
        pTempGeomColl.AddGeometry pSubClip
        Set pTempPoly = New Polygon
        
        Set pTopoOp4 = pTempPoly
        pTopoOp4.ConstructUnion pTempGeomColl
        Set pSubClip = pTempPoly
      End If
      
      Set pTopoOp4 = pSubClip
      pSubClip.SimplifyPreserveFromTo
      pTopoOp4.IsKnownSimple = False
      pTopoOp.Simplify
      
      pSubPolyArray.Add pSubClip
    
      Set pFeature = pFCursor.NextFeature
    Loop
      
    Set pCombinePolygon = UnionGeometries4(pSubPolyArray, 100)
    Set pTopoOp4 = pCombinePolygon
    pTopoOp4.IsKnownSimple = False
    pTopoOp4.Simplify
    pCombinePolygon.SimplifyPreserveFromTo
    
    Set pCombinePolygon.SpatialReference = pGeoSpRef
  End If
  
  Set pTopoOp = pAnalysisPolygon
  Set pCombinePolygon = pTopoOp.Intersect(pCombinePolygon, pCombinePolygon.Dimension)
  
  dblPolygonSumArea = 0
  If booUseSpherical Then
    Set pSpheroid = pGeoSpRef.Datum.Spheroid
    dblSemiMajorAxis = pSpheroid.SemiMajorAxis
    dblSemiMinorAxis = pSpheroid.SemiMinorAxis
    dblPolygonSumArea = dblPolygonSumArea + _
          SphericalPolygonArea2(pCombinePolygon, , , , dblSemiMajorAxis, dblSemiMinorAxis)
  Else
    Set pArea = pCombinePolygon
    dblPolygonSumArea = dblPolygonSumArea + pArea.Area
  End If
  
  ' CALCULATE AREA OF UNCLIPPED POLYGON
  If booUseSpherical Then
    dblFullArea = SphericalPolygonArea2(pAnalysisPolygon, , , , dblSemiMajorAxis, dblSemiMinorAxis)
  Else
    Set pArea = pAnalysisPolygon
    dblFullArea = pArea.Area
  End If
  
  Set pTopoOp = pAOIPolygon
  Dim pClipPolygon As IPolygon
  Set pClipPolygon = pTopoOp.Intersect(pAnalysisPolygon, pAnalysisPolygon.Dimension)
  pClipPolygon.SimplifyPreserveFromTo
  
  ' CALCULATE AREA OF CLIPPED POLYGON
  If booUseSpherical Then
    dblClipArea = SphericalPolygonArea2(pClipPolygon, , , , dblSemiMajorAxis, dblSemiMinorAxis)
  Else
    Set pArea = pClipPolygon
    dblClipArea = pArea.Area
  End If
  
  ' RETURN CLIPPED POLYGON TO ORIGINAL SPATIAL REFERENCE IF NECESSARY
  If booShouldProject Then
    Set pGeom = pClipPolygon
    If Not pTransform Is Nothing Then
      If pTransformDir = esriTransformForward Then
        pGeom.ProjectEx pPolygonToClip.SpatialReference, esriTransformReverse, pTransform, False, 0, 0
      Else
        pGeom.ProjectEx pPolygonToClip.SpatialReference, esriTransformForward, pTransform, False, 0, 0
      End If
    Else
      pGeom.Project pPolygonToClip.SpatialReference
    End If
  End If
  
'  Dim varReturn As Variant
'  varReturn = Array(pClipPolygon, dblClipArea / dblFullArea)
  If dblClipArea = 0 Or dblPolygonSumArea = 0 Then
    dblProportion = 0
  Else
    If dblClipArea > dblPolygonSumArea Then
      dblPolygonSumArea = dblClipArea
    End If
    dblProportion = dblClipArea / dblPolygonSumArea
  End If
  If dblProportion > 1 Then dblProportion = 1
  ReturnClipAndProportion = Array(pClipPolygon, dblFullArea, dblPolygonSumArea, dblClipArea, dblProportion)
 
ClearMemory:
  Set pClone = Nothing
  Set pAnalysisPolygon = Nothing
  Set pGeom = Nothing
  Set pSpheroid = Nothing
  Set pGeoSpRef = Nothing
  Set pTopoOp = Nothing
  Set pRelOp = Nothing
  Set pFCursor = Nothing
  Set pFeature = Nothing
  Set pSubPoly = Nothing
  Set pSubClip = Nothing
  Set pSpFilt = Nothing
  Set pArea = Nothing
  Set pSubPolyArray = Nothing
  Set pCombinePolygon = Nothing
  Set pTopoOp4 = Nothing
  Set pClipPoly = Nothing
  Set pClipEnv = Nothing
  Set pClipPolygon = Nothing


  
End Function

Public Function DistanceBetweenPolygons(booUsingPolygons As Boolean, varSourceObjects_PolysOrDoubleArrays As Variant, _
     Optional dblCloseX1 As Double, Optional dblCloseY1 As Double, _
     Optional dblCloseX2 As Double, Optional dblCloseY2 As Double) As Double

  Dim pPolygon1 As IPolygon
  Dim dblPolyArray1() As Double
  
  Dim pPolygon2 As IPolygon
  Dim dblPolyArray2() As Double
  
  Dim varArrays1() As Variant
  Dim varArrays2() As Variant
  
  Dim lngIndex1 As Long
  Dim lngIndex2 As Long
  Dim lngPointIndex1 As Long
  Dim lngPointIndex2 As Long
      
  If booUsingPolygons Then
    Set pPolygon1 = varSourceObjects_PolysOrDoubleArrays(0)
    Set pPolygon2 = varSourceObjects_PolysOrDoubleArrays(1)
  
  ' RETURNS A VARIANT ARRAY CONTAINING MULTIPLE 2-DIMENSIONAL DOUBLE ARRAYS OF X- AND Y-COORDINATES OF ALL POLYGON RINGS.
  ' EACH DOUBLE ARRAY HAS (VERTEX COUNT) ROWS, WHERE THE LAST ROW/VERTEX IS EQUAL TO THE FIRST ROW/VERTEX.
  ' CODE REMOVES ANY DUPLICATE CONSECUTIVE VERTICES.
  
    varArrays1 = MyGeometricOperations.ReturnPolygonRingsAsDoubleArray(pPolygon1)
    varArrays2 = MyGeometricOperations.ReturnPolygonRingsAsDoubleArray(pPolygon2)
    
  Else
    varArrays1 = varSourceObjects_PolysOrDoubleArrays(0)
    varArrays2 = varSourceObjects_PolysOrDoubleArrays(1)
  End If
  
  Dim dblMinDist As Double
  Dim dblTestDist As Double
  Dim dblSeg1Start(1) As Double
  Dim dblSeg1End(1) As Double
  Dim dblSeg2Start(1) As Double
  Dim dblSeg2End(1) As Double
  Dim dblClose1() As Double
  Dim dblClose2() As Double
  
  dblMinDist = 2 ^ 100
  
  For lngIndex1 = 0 To UBound(varArrays1)
    dblPolyArray1 = varArrays1(lngIndex1)
    
    For lngIndex2 = 0 To UBound(varArrays2)
      dblPolyArray2 = varArrays2(lngIndex2)
      
      For lngPointIndex1 = 0 To UBound(dblPolyArray1, 2) - 1
        For lngPointIndex2 = 0 To UBound(dblPolyArray2, 2) - 1
        
          dblSeg1Start(0) = dblPolyArray1(0, lngPointIndex1)
          dblSeg1Start(1) = dblPolyArray1(1, lngPointIndex1)
          dblSeg1End(0) = dblPolyArray1(0, lngPointIndex1 + 1)
          dblSeg1End(1) = dblPolyArray1(1, lngPointIndex1 + 1)
        
          dblSeg2Start(0) = dblPolyArray2(0, lngPointIndex2)
          dblSeg2Start(1) = dblPolyArray2(1, lngPointIndex2)
          dblSeg2End(0) = dblPolyArray2(0, lngPointIndex2 + 1)
          dblSeg2End(1) = dblPolyArray2(1, lngPointIndex2 + 1)
        
          dblTestDist = SquaredDistanceBetweenSegments(dblSeg1Start, dblSeg1End, dblSeg2Start, dblSeg2End, dblClose1, dblClose2)
          
          If dblTestDist < dblMinDist Then
            dblMinDist = dblTestDist
            dblCloseX1 = dblClose1(0)
            dblCloseY1 = dblClose1(1)
            dblCloseX2 = dblClose2(0)
            dblCloseY2 = dblClose2(1)
          End If
        Next lngPointIndex2
      Next lngPointIndex1
    Next lngIndex2
  Next lngIndex1
  
  
  ' REMEMBER RUNNING DISTANCES WERE SQUARED
  DistanceBetweenPolygons = Sqr(dblMinDist)
  
' SAMPLE ====================================
'  Dim pMxDoc As IMxDocument
'  Set pMxDoc = ThisDocument
'
'  Dim strFolder As String
'  strFolder = "D:\arcGIS_stuff\consultation\Margaret_Moore\Modified_Data_August_4_2019\Combined_by_Quadrat.gdb"
'
'  Dim strFClassName As String
'  strFClassName = "Cover_All"
'
'  Dim lngIndex As Long
'  Dim pWS As IFeatureWorkspace
'  Dim pWSFact As IWorkspaceFactory
'  Dim pFClass As IFeatureClass
'  Dim pFeature As IFeature
'  Dim pPolygon As IPolygon
'  Dim pPolygon2 As IPolygon
'  Dim pGeoDataset As IGeoDataset
'  Dim pSpRef As ISpatialReference
'  Dim pSpRefRes As ISpatialReferenceResolution
'  Dim dblDist2 As Double
'
'  Dim pRelOp As IRelationalOperator
'  Dim pProxOp As IProximityOperator
'  Dim pTopoOp As ITopologicalOperator3
'
'  Set pWSFact = New FileGDBWorkspaceFactory
'  Set pWS = pWSFact.OpenFromFile(strFolder, 0)
'  Debug.Print "--------------------"
'  Debug.Print strFClassName
'  Set pFClass = pWS.OpenFeatureClass(strFClassName)
'
'  Set pGeoDataset = pFClass
'  Set pSpRef = pGeoDataset.SpatialReference
'  Set pSpRefRes = pSpRef
'
'  Debug.Print pSpRef.Name
'  Debug.Print "Resolution [True] = " & Format(pSpRefRes.XYResolution(True), "0.000000000000")
'  Debug.Print "Resolution [False] = " & Format(pSpRefRes.XYResolution(False), "0.000000000000")
'
'  Set pFeature = pFClass.GetFeature(15189)
'  Set pPolygon = pFeature.ShapeCopy
'  Set pFeature = pFClass.GetFeature(15190)
'  Set pPolygon2 = pFeature.ShapeCopy
'
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon, "Delete_Me"
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPolygon2, "Delete_Me"
'
'  Set pRelOp = pPolygon
'  Debug.Print "Is Disjoint = " & CStr(pRelOp.Disjoint(pPolygon2))
'  Debug.Print "Touches = " & CStr(pRelOp.Touches(pPolygon2))
'
'  Set pProxOp = pPolygon
'  Debug.Print "Prox Op Distance = " & CStr(pProxOp.ReturnDistance(pPolygon2))
'
'  Dim dblCloseX1 As Double
'  Dim dblCloseY1 As Double
'  Dim dblCloseX2 As Double
'  Dim dblCloseY2 As Double
'
'  dblDist2 = MyGeometricOperations.DistanceBetweenPolygons(True, Array(pPolygon, pPolygon2), _
'      dblCloseX1, dblCloseY1, dblCloseX2, dblCloseY2)
'
'  Dim pPoint1 As IPoint
'  Dim pPoint2 As IPoint
'  Set pPoint1 = New Point
'  Set pPoint2 = New Point
'  Set pPoint1.SpatialReference = pPolygon.SpatialReference
'  Set pPoint2.SpatialReference = pPolygon.SpatialReference
'  pPoint1.PutCoords dblCloseX1, dblCloseY1
'  pPoint2.PutCoords dblCloseX2, dblCloseY2
'
'  Debug.Print "Manual Distance = " & Format(dblDist2 * 1000#, "0.000") & " mm"
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint1, "Delete_Me"
'  MyGeneralOperations.Graphic_MakeFromGeometry pMxDoc, pPoint2, "Delete_Me"
'
'
'ClearMemory:
'  Set pMxDoc = Nothing
'  Set pWS = Nothing
'  Set pWSFact = Nothing
'  Set pFClass = Nothing
'  Set pFeature = Nothing
'  Set pPolygon = Nothing
'  Set pPolygon2 = Nothing
  
ClearMemory:
  Set pPolygon1 = Nothing
  Erase dblPolyArray1
  Set pPolygon2 = Nothing
  Erase dblPolyArray2
  Erase varArrays1
  Erase varArrays2
  Erase dblSeg1Start
  Erase dblSeg1End
  Erase dblSeg2Start
  Erase dblSeg2End
  Erase dblClose1
  Erase dblClose2



End Function
    

