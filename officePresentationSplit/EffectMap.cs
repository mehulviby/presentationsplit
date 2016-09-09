using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Diagnostics;

namespace officePresentationSplit
{
    internal class EffectMap
    {
        internal static bool isPathEffect(Effect e)
        {
            // TODO: On Error GoTo Warning!!!: The statement is not translatable 
            bool isPathEffect = false;
            try
            {
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathCrescentMoon);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSquare);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathTrapezoid);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathHeart);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathOctagon);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPath4PointStar);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPath5PointStar);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPath6PointStar);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPath8PointStar);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathFootball);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathEqualTriangle);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathParallelogram);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathPentagon);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathTeardrop);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathPointyStar);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathCurvedSquare);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathCurvedX);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathVerticalFigure8);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathCurvyStar);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathLoopdeLoop);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathBuzzsaw);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathHorizontalFigure8);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathPeanut);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathFigure8Four);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathNeutron);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSwoosh);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathBean);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathPlus);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathInvertedTriangle);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathInvertedSquare);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathLeft);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathTurnRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathArcDown);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathZigzag);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSCurve2);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSineWave);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathBounceLeft);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathDown);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathTurnUp);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathArcUp);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathHeartbeat);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSpiralRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathWave);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathCurvyLeft);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathDiagonalDownRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathTurnDown);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathArcLeft);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathFunnel);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSpring);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathBounceRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSpiralLeft);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathDiagonalUpRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathTurnUpRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathArcRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathSCurve1);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathDecayingWave);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathCurvyRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathStairsDown);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathUp);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectPathRight);
                isPathEffect = (isPathEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectCustom);
                return isPathEffect;
            }
            catch (Exception)
            {
                Trace.WriteLine("Path Effect :::::");
                return false;
            }
        }
        internal static bool isEmphasisEffect(Effect e)
        {
            // TODO: On Error GoTo Warning!!!: The statement is not translatable 
            bool isEmphasisEffect = false;
            try
            {
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectGrowShrink);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectChangeFontColor);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectChangeFillColor);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectChangeFontStyle);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectTransparency);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectChangeFont);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectChangeLineColor);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectChangeFontSize);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectSpin);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectDesaturate);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectColorWave);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectComplementaryColor2);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectVerticalGrow);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectLighten);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectColorBlend);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectBrushOnUnderline);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectBrushOnColor);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectComplementaryColor);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectContrastingColor);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectBoldFlash);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectFlashBulb);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectDarken);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectGrowWithColor);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectTeeter);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectFlicker);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectBoldReveal);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectWave);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectStyleEmphasis);
                isEmphasisEffect = (isEmphasisEffect) | (e.EffectType == MsoAnimEffect.msoAnimEffectBlast);

                isEmphasisEffect = (isEmphasisEffect | isPathEffect(e));
                //  If isEmphasisEffect is true at this point, then I have
                //  an emphasis or motion effect. But let's really make sure it is not
                //  an entry/exit effect.

                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectAppear));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFly));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectBlinds));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectBox));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectCheckerboard));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectCircle));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectCrawl));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectDiamond));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectDissolve));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFade));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFlashOnce));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectPeek));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectPlus));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectRandomBars));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectSpiral));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectSplit));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectStretch));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectStrips));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectSwivel));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectWedge));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectWheel));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectWipe));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectZoom));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectRandomEffects));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectBoomerang));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectBounce));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectColorReveal));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectCredits));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectEaseIn));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFloat));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectGrowAndTurn));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectLightSpeed));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectPinwheel));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectRiseUp));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectSwish));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectThinLine));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectUnfold));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectWhip));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectAscend));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectCenterRevolve));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFadedSwivel));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectDescend));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectSling));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectSpinner));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectStretchy));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectZip));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectArcUp));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFadedZoom));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectGlide));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectExpand));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFlip));
                isEmphasisEffect = (isEmphasisEffect & (e.EffectType != MsoAnimEffect.msoAnimEffectFold));
                return isEmphasisEffect;
            }
            catch (Exception)
            {
                Trace.WriteLine("Emphasis :::::");
                return true;
            }
        }
    }
}
