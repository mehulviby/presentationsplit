using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Threading;
using System.IO;
using System.Diagnostics;
using System.Collections;
using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace officePresentationSplit
{
    public partial class MyAuthPipeline1
    {
        private Application _pptApplication;
        private Presentation duplicatePresentation;
        
        public bool splitMouseTriggered;
        public int slide_number;
        public const int maxProgressWidth = 324;

        private void MyAuthPipeline1_Load(object sender, RibbonUIEventArgs e)
        {
            Trace.WriteLine("Ribbon load");
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            this._pptApplication = Globals.ThisAddIn.Application;
            this.splitMouseTriggered = true;

            var sourcePath = Globals.ThisAddIn.Application.ActivePresentation.FullName;
            var fileName = Path.GetFileNameWithoutExtension(sourcePath);
            var ext = Path.GetExtension(sourcePath);

            var destination = @"D:\" + "_testing" + ext;
            File.Copy(sourcePath, destination, true);
            this.duplicatePresentation = this._pptApplication.Presentations.Open(destination, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoFalse);
            foreach (Slide slide in this.duplicatePresentation.Slides)
            {
                int count = slide.TimeLine.MainSequence.Count;
                slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectNone;
                slide.SlideShowTransition.AdvanceTime = 0;
                slide.SlideShowTransition.Duration = 0;
                foreach (Effect effect in slide.TimeLine.MainSequence)
                {
                    var exitProperty = effect.Exit;
                    effect.EffectType = MsoAnimEffect.msoAnimEffectAppear;
                    effect.Exit = exitProperty;

                }
            }
            this.duplicatePresentation.Save();
            this.PPspliT_main();
            this.duplicatePresentation.Save();
            this.duplicatePresentation = this._pptApplication.Presentations.Open(destination, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }
        private void PPspliT_main()
        {
            try
            {
                Sequence slide_timeline = null;
                int tot_slides = 1;
                this.slide_number = 1;
                tot_slides = this.duplicatePresentation.Slides.Count;

                this.bakeSlideNumbers(this.slide_number, tot_slides, maxProgressWidth);

                int orig_tot_slides = tot_slides;
                int actual_slide = this.slide_number;
                while (actual_slide <= tot_slides)
                {
                    bool additional_slide_present = false;
                    bool alreadyPurged = false;
                    if (this.duplicatePresentation.Slides[this.slide_number].TimeLine.MainSequence.Count > 0)
                    {
                        this.copyShapeIds(this.duplicatePresentation.Slides[this.slide_number]);
                        bool cont = (duplicatePresentation.Slides[this.slide_number].TimeLine.MainSequence[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerWithPrevious ||
                                    duplicatePresentation.Slides[this.slide_number].TimeLine.MainSequence[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        if (cont)
                        {
                            /* Keep a copy of the original slide, which I will use to track the animation
                            sequence. I always proceed in this way: I carry the original slide
                            unaltered and grab the list of effects to be applied from it, while
                            shapes are actually modified on copies of that original slide */
                            this.duplicatePresentation.Slides[this.slide_number].Duplicate();
                            //Remember to remove the duplicated slide later on
                            additional_slide_present = true;
                            slide_timeline = this.duplicatePresentation.Slides[this.slide_number + 1].TimeLine.MainSequence;
                            //Remove all the shapes that will appear after a future entry effect

                            this.RemoveFutureShapes(this.duplicatePresentation.Slides[this.slide_number], false);
                            this.RemoveEffects(this.duplicatePresentation.Slides[this.slide_number]);
                            alreadyPurged = true;
                        }
                        while (cont)
                        {
                            //Actually, there are animations that start without a click
                            this.applyEffect(this.duplicatePresentation.Slides[this.slide_number], this.duplicatePresentation.Slides[this.slide_number + 1]);
                            // Some effects have disappeared: check whether I still have
                            // effects that start without a click
                            if (slide_timeline.Count == 0)
                                cont = false;
                            else
                                cont = (slide_timeline[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerWithPrevious)
                                    || (slide_timeline[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerAfterPrevious);
                        }
                        if (additional_slide_present)
                        {
                            this.matchZOrder(this.duplicatePresentation.Slides[this.slide_number], this.duplicatePresentation.Slides[this.slide_number + 1]);
                        }
                    }
                    else
                        actual_slide++;

                    int tot_anims;
                    if(additional_slide_present)
                        tot_anims = this.duplicatePresentation.Slides[this.slide_number + 1].TimeLine.MainSequence.Count;
                    else
                        tot_anims = this.duplicatePresentation.Slides[this.slide_number].TimeLine.MainSequence.Count;

                    if(tot_anims > 0)
                    {
                        int processed_anims = 0;
                        if(!alreadyPurged)
                        {
                            this.duplicatePresentation.Slides[this.slide_number].Duplicate();
                            this.RemoveFutureShapes(this.duplicatePresentation.Slides[this.slide_number], false);
                            this.RemoveEffects(this.duplicatePresentation.Slides[this.slide_number]);
                            alreadyPurged = true;
                        }
                        this.duplicatePresentation.Slides[this.slide_number].Duplicate();
                        this.slide_number++;

                        while (this.duplicatePresentation.Slides[slide_number + 1].TimeLine.MainSequence.Count > 0)
	                    {
	                        bool cont = true;
                            while(cont)
                            {
                                int addedEffects = this.applyEffect(this.duplicatePresentation.Slides[this.slide_number], this.duplicatePresentation.Slides[this.slide_number + 1]);
                                slide_timeline = this.duplicatePresentation.Slides[this.slide_number + 1].TimeLine.MainSequence;
                                if(slide_timeline.Count == 0)
                                    cont = false;
                                else
                                    cont = ((slide_timeline[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerWithPrevious)
                                    || (slide_timeline[1].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerAfterPrevious));
                                processed_anims += 1 - addedEffects;
                                int anims_percentage = Convert.ToInt16(processed_anims / tot_anims * 100);
                            }
                            this.matchZOrder(this.duplicatePresentation.Slides[this.slide_number], this.duplicatePresentation.Slides[this.slide_number + 1]);
                            if (slide_timeline.Count > 0)
                            {
                                this.duplicatePresentation.Slides[this.slide_number].Duplicate();
                                this.RemoveEffects(this.duplicatePresentation.Slides[this.slide_number]);
                                this.slide_number++;
                            }
	                    }
                        this.duplicatePresentation.Slides[this.slide_number + 1].Delete();
                        additional_slide_present = false;
                        this.RemoveEffects(this.duplicatePresentation.Slides[this.slide_number]);
                        actual_slide++;
                    }
                    this.slide_number++;
                }
            }
            catch (Exception)
            {}
        }
        // 
        //  This subroutine applies the ZOrder (depth) of shapes in s2 to shapes in s1.
        //  Corresponding shapes in s1 and in s2 are different objects, therefore, in order
        //  to be matched, shape IDs must have been copied to the AlternativeText in advance
        //  by using the copyShapeIds subroutine.
        //  Note: the algorithm used to sort shapes in s2 by increasing ZOrder could be
        //  improved.
        // 
        private void matchZOrder(Slide s1, Slide s2)
        {
            //ArrayList list = new ArrayList();
            int zThreshold = 0;
            int j = 1;
            string minZshapeId = null;
            try
            {
                for (int i = 1; (i <= s2.Shapes.Count); i++)
                {
                    int minZ = 65536;
                    //  Find shape in s2 with minimum ZOrder greater than zThreshold
                    foreach (PowerPoint.Shape sh2 in s2.Shapes)
                    {
                        //  Inequalities are strict because there should be no
                        //  two shapes with the same ZOrder
                        if (((sh2.ZOrderPosition < minZ) && (sh2.ZOrderPosition > zThreshold)))
                        {
                            minZ = sh2.ZOrderPosition;
                            minZshapeId = sh2.AlternativeText;
                        }
                    }
                    zThreshold = minZ;
                    int shapeIdInS1 = this.findShape(s1, minZshapeId);
                    if (shapeIdInS1 > 0)
                    {
                        //  The same shape exists also in s1: add the shape to the array of sorted shapes
                        s1.Shapes[shapeIdInS1].ZOrder(MsoZOrderCmd.msoBringToFront);
                    }
                }

                //  Bring to front shapes in s1 by increasing values of ZOrder
                //for (int i = 1; i <= (j - 1); i++)
                //{
                //    sortedShapes[i].ZOrder(MsoZOrderCmd.msoBringToFront);
                //}
            }
            catch (Exception exp)
            {
                Trace.WriteLine("Message : " + exp.Message);
            }
        }
        private void bakeSlideNumbers(int start_index, int end_index, int maxProgressWidth)
        {
            for (int i = start_index; i <= end_index; i++)
            {
                foreach (PowerPoint.Shape shape in this.duplicatePresentation.Slides[i].Shapes)
                {
                    if (shape.Type == MsoShapeType.msoPlaceholder)
                    {
                        var placeholder = shape.PlaceholderFormat;
                        if(placeholder.Type == PpPlaceholderType.ppPlaceholderSlideNumber 
                            || placeholder.Type == PpPlaceholderType.ppPlaceholderDate
                            || placeholder.Type == PpPlaceholderType.ppPlaceholderFooter)
                        {
                            for (int c = 1; c <= shape.TextFrame.TextRange.Characters().Count; c++)
                            {
                                shape.TextFrame.TextRange.Characters(c).Text = shape.TextFrame.TextRange.Characters(c).Text;
                            }
                        }
                    }
                }
            }
        }
        private void copyShapeIds(Slide s){
            foreach (PowerPoint.Shape sh in s.Shapes)
	        {
		        sh.AlternativeText = sh.Id.ToString();
	        }
        }
        private void RemoveFutureShapes(Slide slide, bool textParagraphEffectsOnly)
        {
            Sequence slide_timeline = slide.TimeLine.MainSequence;
            int i = 1, start_deleting_at = 0;
            while (i <= slide_timeline.Count && start_deleting_at == 0)
            {
                if(slide_timeline[1].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerAfterPrevious
                    && slide_timeline[1].Timing.TriggerType != MsoAnimTriggerType.msoAnimTriggerWithPrevious)
                {
                    /*Start deleting shapes from the next mouse-triggered event.
                    Any preceding shapes will be deleted when their effects
                    are individually considered*/
                    start_deleting_at = i;
                }
                i++;
            }
            if (start_deleting_at > 0)
            {
                for (i = start_deleting_at; i <= slide.TimeLine.MainSequence.Count; i++)
                {
                    if (i > slide.TimeLine.MainSequence.Count)
                        break;
                    int delete_shape_idx = -1;
                    if(!(slide_timeline[i].Exit == MsoTriState.msoTrue))
                    {
                        delete_shape_idx = i;
                    }
                    int parI = this.getEffectParagraph(slide_timeline[i]);
                    Trace.WriteLine(" J " + (i - 1) + " : start_deleting_at : " + start_deleting_at);
                    for (int j = i - 1; j >= start_deleting_at; j--)
                    {
                        try
                        {
                            if (slide_timeline[i].Shape == slide_timeline[j].Shape && (!(slide_timeline[j].Exit == MsoTriState.msoTrue)))
                            {
                                /*Probably we need to abort deletion: there may
                                be an exit/emphasis effect for the same shape before the entry effect.
                                In that case, this means that the shape must be visible at the
                                beginning. However, first we need to check if this is a paragraph
                                effect and, in that case, if the exit/emphasisfor 
                                effect applies to the very same paragraph.*/
                                int parJ = this.getEffectParagraph(slide_timeline[j]);
                                if (parI == parJ)
                                {
                                    /*Either none of the effects is a paragraph effect (in which
                                    case the match is ok because both effects work on the same shape)
                                    or both effects are paragraph effects and work on the same paragraph
                                    (in which case the match is still ok because they affect the
                                    same graphical element). If the match is ok, then deletion
                                    must be aborted.*/
                                    delete_shape_idx = -1;
                                }
                            } 
                        }
                        catch (Exception)
                        {
                        }
                    }
                    if (delete_shape_idx > 0)
                    {
                        /*Delete shapes for which a following entry effect exists.
                        Restrict deletion to text paragraphs only if instructed to
                        do so.*/
                        if (parI > 0 || !textParagraphEffectsOnly)
                        {
                            /*Pay attention, because shape deletion (not paragraph deletion)
                            causes animation effects to disappear from the timeline, so we
                            need to decrease i in order to keep in sync with the currently
                            processed effect.
                            In general, deletion of a shape may cause several preceding
                            effects to also disappear: here we count how many in order to
                            understand how many positions should i go backward (note that
                            future effects for the same shapes should not be counted, because
                            they will safely disappear from the timeline without the need
                            to realign the value of i).*/
                            int prevEffectsForThisShape = 0;
                            for (int k = 1; k <= i; k++)
                            {
                                if (slide_timeline[k].Shape == slide_timeline[i].Shape)
                                {
                                    prevEffectsForThisShape++;
                                }
                            }
                            /*Assertion: at the end of the above iteration, prevEffectsForThisShape
                            should always be >0 (because at least the i'th effect affects that
                            shape)*/
                            if(this.deleteShape(slide_timeline[i].Shape, slide_timeline, delete_shape_idx))
                            {
                                i = i - prevEffectsForThisShape;
                            }
                        }
                    }
                }
            }
        }
        // 
        //  This function takes an effect as argument. If the
        //  effect is applied to a text paragraph, it returns the
        //  index of that text paragraph (in its container shape).
        //  Otherwise, it returns -1.
        // 
        private int getEffectParagraph(Effect e)
        {
            int paragraph_idx = -1;
            try
            {
                paragraph_idx = e.Paragraph;
            }
            catch (Exception)
            {
            }
            return paragraph_idx;
        }
        private bool deleteShape(PowerPoint.Shape shape, Sequence theTimeline, int effectId)
        {
            MsoAnimTriggerType animType = MsoAnimTriggerType.msoAnimTriggerNone;
            bool deleteShape;
            int theParagraph = this.getEffectParagraph(theTimeline[effectId]);
            if (theParagraph > 0)
            {
                int oldCount = theTimeline.Count;
                if (oldCount > effectId)
                {
                    animType = theTimeline[effectId + 1].Timing.TriggerType;
                }
                this.clearParagraph(shape, theParagraph);
                if (theTimeline.Count < oldCount)
                {
                    /*The removed paragraph was not the last one in the shape, and therefore
                    the effect has been automatically removed. Restore the trigger
                    type if required*/
                    if (theTimeline.Count >= effectId)
                    {
                        //Restore the trigger type
                        theTimeline[effectId].Timing.TriggerType = animType;
                    }
                    deleteShape = true;
                }
                else
                {
                    // The removed paragraph was the last one in the shape, therefore
                    // the effect is still there.
                    deleteShape = false;
                }
            }
            else{
                // whole shape effect
                shape.Delete();
                deleteShape = true;
            }
            return deleteShape;
        }
        private void clearParagraph(PowerPoint.Shape shape, int par)
        {
            int i;
            if (shape.TextFrame.TextRange.Paragraphs(par).Lines().Count > 1)
            {
                 /*This is a word wrapped or multi-line paragraph: turn every
                word wrap into a real new line. This is required because the
                paragraph contents will be soon replaced with spaces, which
                have a different width than original characters, can therefore
                mess up word wrapping, hence the number of lines of this paragraph,
                hence the rendering of any following paragraphs.*/
                for (i = 2; i <= shape.TextFrame.TextRange.Paragraphs(par).Lines().Count; i++)
                {
                    TextRange textRange = shape.TextFrame.TextRange.Paragraphs(par).Lines(i - 1).Characters(shape.TextFrame.TextRange.Paragraphs(par).Lines(i - 1).Characters().Count);
                    int asciiCode = this.ReturnAsciiCode(textRange.Text);
                    if ((asciiCode != 11) && (asciiCode != 13))
                    {
                        shape.TextFrame.TextRange.Paragraphs(par).Lines(i).Characters(1).InsertBefore(((char)11).ToString());
                    }
                }
            }
            TextRange p = shape.TextFrame.TextRange.Paragraphs(par);
            i = 1;
            while (i <= p.Characters().Count)
            {
                /*Replace paragraph contents with spaces. This is the best and
                most compatible way I found to "hide" a paragraph while keeping
                its original space occupied.*/
                int asciiCode = this.ReturnAsciiCode(p.Characters(i).Text);
                if (asciiCode != 13 && asciiCode != 11)
                {
                    p.Characters(i).Text = " ";
                }
                i++;
            }
            //Set bullet symbol too to " " (32 is the Unicode value)
            p.ParagraphFormat.Bullet.Character = 32;
        }
        private void RemoveEffects(Slide slide)
        {
            for (int i = 1; i <= slide.TimeLine.MainSequence.Count; i++)
            {
                slide.TimeLine.MainSequence[1].Delete();
            }
            slide.SlideShowTransition.EntryEffect = PpEntryEffect.ppEffectAppear;
        }
        private int applyEffect(Slide slide, Slide seq_slide)
        {
            Effect current_effect;
            PowerPoint.Shape shape;
            current_effect = seq_slide.TimeLine.MainSequence[1];
            shape = current_effect.Shape;
            int applyEffect = 0;
            if (current_effect.EffectInformation.AfterEffect == MsoAnimAfterEffect.msoAnimAfterEffectHide)
            {
                // This effect is set for hiding the shape after the animation, so it
                // must be treated equivalently to an exit effect: simply delete the shape
                if(this.findShape(slide, shape.AlternativeText) > 0)
                {
                    this.deleteShape(slide.Shapes[this.findShape(slide, shape.AlternativeText)], seq_slide.TimeLine.MainSequence, 1);
                }
                current_effect.Delete();
            }
            else
            {
                if (current_effect.EffectInformation.AfterEffect == MsoAnimAfterEffect.msoAnimAfterEffectHideOnNextClick)
                {
                    // This effect is set for hiding after the next click:
                    // insert a new exit animation that will be processed in the following
                    bool found = false;
                    Sequence tl = seq_slide.TimeLine.MainSequence;
                    int i;
                    for (i = 2; i <= tl.Count; i++)
                    {
                        if (tl[i].Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                        {
                            tl.AddEffect(current_effect.Shape, MsoAnimEffect.msoAnimEffectDissolve, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerWithPrevious);
                            /*Best thing would be to insert the exit effect right after the next click-triggered
                            effect, but this is not possible, guess why, due to a PowerPoint bug which causes
                            the Index argument of AddEffect to be handled unpredictably. So, we need to work this
                            around by inserting the effect at the end of the sequence and, only afterwards,
                            move it to the right location.*/
                            tl[tl.Count].MoveTo(i + 1);
                            tl[i + 1].Exit = MsoTriState.msoTrue;
                            found = true;
                            break;
                        }
                    }
                    if(!found)
                    {
                        tl.AddEffect(current_effect.Shape, MsoAnimEffect.msoAnimEffectDissolve, MsoAnimateByLevel.msoAnimateLevelNone, MsoAnimTriggerType.msoAnimTriggerOnPageClick, i);
                        tl[i].Exit = MsoTriState.msoTrue;
                    }
                     /*This is the only case when the applyEffect function adds an animation effect to the
                    sequence: here we notify the calling routine about the fact that the animation sequence
                    has lengthened.*/
                    applyEffect = 1;
                }
                if (current_effect.Timing.RewindAtEnd == MsoTriState.msoTrue)
                    current_effect.Delete();
                else
                {
                    if (current_effect.Exit == MsoTriState.msoTrue)
                    {
                        //This is an exit effect: simply delete the shape (or the text
                        //paragraph) from the next slide
                        if (this.findShape(slide, shape.AlternativeText) > 0)
                        {
                            this.deleteShape(slide.Shapes[this.findShape(slide, shape.AlternativeText)], seq_slide.TimeLine.MainSequence, 1);
                        }
                        current_effect.Delete();
                    }
                    else
                    {
                        if (!(this.findShape(slide, shape.AlternativeText) > 0))
                        {
                            /*' The shape is not already present
                            ' Invoke purgeEffects to clear any subsequent entry
                            ' effects, which may interfere
                            ' with calls to purgeFutureShapes below in this same
                            ' subroutine.
                            ' (note that these subsequent calls may happen when
                            ' in the same slide multiple objects appear simultaneously,
                            ' and therefore applyEffect is invoked multiple times).*/
                            shape.Copy();
                            this.RemoveEffects(slide);
                            slide.Shapes.Paste();
                            PowerPoint.Shape newShape = slide.Shapes[this.findShape(slide, shape.AlternativeText)];
                            newShape.Left = shape.Left;
                            newShape.Top = shape.Top;
                            int par = -1;
                            try 
	                        {	        
		                        par = current_effect.Paragraph;
	                        }
	                        catch (Exception)
	                        {
	                        }
                            if(par > 0 )
                            {
                                for (int parIdx = 1 ; parIdx <= newShape.TextFrame.TextRange.Paragraphs().Count ; parIdx++)
			                    {
			                        if (parIdx != par)
                                    {
                                        bool foundEntryAnim = false;
                                        for (int k = 1; k <= seq_slide.TimeLine.MainSequence.Count; k++) {
                                            if ((seq_slide.TimeLine.MainSequence[k].Shape == shape) && !(seq_slide.TimeLine.MainSequence[k].Exit == MsoTriState.msoTrue))
                                            {
                                                try 
	                                            {	        
                                                    if ((seq_slide.TimeLine.MainSequence[k].Paragraph == parIdx)) {
                                                        foundEntryAnim = true;
                                                    }
	                                            }
	                                            catch (Exception)
	                                            {
	                                            }
                                            }
                                        }
            
                                        if (foundEntryAnim) {
                                            this.clearParagraph(slide.Shapes[this.findShape(slide, shape.AlternativeText)],parIdx);
                                        }
            
                                    }
			                    }
                            }
                            //  Sometimes text auto-fitting does not seem to act
                            //  properly: this is an attempt to "awaken" it by
                            //  notifying of a change in the shape size
                            newShape.Width = shape.Width;
                            newShape.Height = shape.Height;
                            //  Now we have pasted the shape. Note that we paste
                            //  only one shape at a time, therefore it should carry
                            //  with itself its own entry effect. There is one
                            //  exception: a single text box shape may be associated with
                            //  several subsequent entry effects, that correspond
                            //  to paragraphs in the text appearing one after the
                            //  other (and after the text box itself has appeared).
                            //  We should get rid of paragraphs that are supposed
                            //  to appear later on, and this is why we call purgeFutureShapes
                            //  also here. Note that we should remove the entry effect
                            //  for the shape we have just added before invoking
                            //  purgeFutureShapes, or the shape itself will be
                            //  deleted!
                            slide.TimeLine.MainSequence[1].Delete();
                            this.RemoveFutureShapes(slide, true);
                        }
                        else
                        {
                            //  The shape is already present: I only need to add a
                            //  paragraph to it, if required.
                            int par = -1;
                            //  The following assignment may raise an error for missing
                            //  Paragraph property
                            try 
	                        {	        
                                par = current_effect.Paragraph;
	                        }
	                        catch (Exception)
	                        {
	                        }
                            if ((par > 0)) {
                                PowerPoint.Shape newShape = slide.Shapes[this.findShape(slide, shape.AlternativeText)];
                                this.copyParagraph(newShape.TextFrame.TextRange.Paragraphs(par), shape.TextFrame.TextRange.Paragraphs(par));
                                //  Attempt to preserve indentations and margins (these are not
                                //  part of paragraph information, but rather of a Ruler object).
                                //  In principle, the number of ruler levels (i.e., possible
                                //  indentation levels) is fixed. However, according to the documentation
                                //  it should be 5 whereas in practice I have seen cases where it
                                //  counts up to 9. To stay on the safe side, the number of
                                //  ruler levels here is parametric.
                                for (int ruler_level = 1; ruler_level <= shape.TextFrame.Ruler.Levels.Count; ruler_level++) {
                                    //  For some obscure reasons, out-of-range margins are sometimes
                                    //  returned (for example, corresponding to the smallest possible
                                    //  value in a Long variable). In this case, it's better to
                                    //  refrain from copying the margin value, or an error would be
                                    //  raised.
                                    if (Math.Abs(shape.TextFrame.Ruler.Levels[ruler_level].FirstMargin) < 10000000)
                                    {
                                        newShape.TextFrame.Ruler.Levels[ruler_level].FirstMargin = shape.TextFrame.Ruler.Levels[ruler_level].FirstMargin;
                                    }
        
                                    if (Math.Abs(shape.TextFrame.Ruler.Levels[ruler_level].LeftMargin) < 10000000)
                                    {
                                        newShape.TextFrame.Ruler.Levels[ruler_level].LeftMargin = shape.TextFrame.Ruler.Levels[ruler_level].LeftMargin;
                                    }
        
                                }
    
                                //  Sometimes text auto-fitting does not seem to act
                                //  properly: this is an attempt to "awaken" it by
                                //  notifying of a change in the shape size
                                newShape.Width = shape.Width;
                                newShape.Height = shape.Height;
                            }
                        }
                        current_effect.Delete();
                    }
                }
            }
            return applyEffect;
        }
        // 
        //  Copies the contents of p2 into p1.
        //  This is used to restore a previously hidden paragraph.
        // 
        private void copyParagraph(TextRange p1, TextRange p2)
        {
            int asciiCode = this.ReturnAsciiCode(p2.Characters(p2.Characters().Count).Text);
            bool newLineInserted = false;
            if (asciiCode != 13)
            {
                //  This paragraph does not end with a new line (most
                //  likely because it is the last paragraph in the text
                //  frame). Here I add it because I can get all the
                //  formatting attributes of a paragraph only if it
                //  ends with a new line (this is PowerPoint magic...)
                p2.Characters().InsertAfter("\r");
                newLineInserted = true;
            }

            //  Apply contents and formatting from the original paragraph
            p2.Copy();
            //  It seems that the following 3 assignments, applied *before* pasting
            //  the paragraph, reduce the number of cases in which bullet symbols
            //  are lost. The reason why this happens is completely obscure to me, but
            //  repeating the assignment *after* pasting (where this should happen)
            //  seems to be harmless.
            p1.ParagraphFormat.SpaceAfter = p2.ParagraphFormat.SpaceAfter;
            p1.ParagraphFormat.SpaceBefore = p2.ParagraphFormat.SpaceBefore;
            p1.ParagraphFormat.SpaceWithin = p2.ParagraphFormat.SpaceWithin;

            p1.Paste();

            p1.IndentLevel = p2.IndentLevel;
            p1.ParagraphFormat.SpaceAfter = p2.ParagraphFormat.SpaceAfter;
            p1.ParagraphFormat.SpaceBefore = p2.ParagraphFormat.SpaceBefore;
            p1.ParagraphFormat.SpaceWithin = p2.ParagraphFormat.SpaceWithin;
            //  Restore bullet formatting. Since there seems to be no
            //  way to get the currently used image for a bullet, care
            //  must be taken in updating the bullet attributes only if
            //  required, otherwise the applied image may be messed up
            //  and I may be unable to restore it.
            if ((p1.ParagraphFormat.Bullet.Type != p2.ParagraphFormat.Bullet.Type))
            {
                p1.ParagraphFormat.Bullet.Type = p2.ParagraphFormat.Bullet.Type;
            }

            if ((p2.ParagraphFormat.Bullet.Type == PpBulletType.ppBulletUnnumbered)
                        && (p1.ParagraphFormat.Bullet.Character != p2.ParagraphFormat.Bullet.Character))
            {
                p1.ParagraphFormat.Bullet.Character = p2.ParagraphFormat.Bullet.Character;
                this.copyFontAttributes(p1.ParagraphFormat.Bullet.Font, p2.ParagraphFormat.Bullet.Font);
            }

            if ((p2.ParagraphFormat.Bullet.Type == PpBulletType.ppBulletNumbered)
                        && (p1.ParagraphFormat.Bullet.StartValue != p2.ParagraphFormat.Bullet.StartValue))
            {
                p1.ParagraphFormat.Bullet.StartValue = p2.ParagraphFormat.Bullet.StartValue;
            }

            if ((p2.ParagraphFormat.Bullet.Type == PpBulletType.ppBulletNumbered)
                        && (p1.ParagraphFormat.Bullet.Style != p2.ParagraphFormat.Bullet.Style))
            {
                p1.ParagraphFormat.Bullet.Style = p2.ParagraphFormat.Bullet.Style;
            }

            //  It's not over yet.
            //  Paste often acts in an "intelligent" way, by cutting away
            //  apparently useless spaces and other stuff. Here I need a
            //  really accurate paste, which preserves all the characters,
            //  therefore I overwrite (or enrich) the set of previously
            //  pasted characters. Overwriting the characters one by one
            //  ensures that the rest of formatting is left untouched, but
            //  here I may still be adding new text (e.g., new spaces), to
            //  which formatting must be applied. This is the reason of the
            //  call to copyFontAttributes.
            for (int i = 1; (i <= p2.Characters().Count); i++)
            {
                p1.Characters(i).Text = p2.Characters(i).Text;
                this.copyFontAttributes(p1.Characters(i).Font, p2.Characters(i).Font);
            }

            //  Remove any previously inserted new line characters
            if (newLineInserted)
            {
                p1.Characters(p1.Text.Length).Delete();
            }

        }
        // 
        //  Copies fundamental font attributes from f2 to f1.
        // 
        private void copyFontAttributes(Font f1, Font f2) {
            f1.Name = f2.Name;
            f1.Size = f2.Size;
            f1.Bold = f2.Bold;
            f1.Italic = f2.Italic;
            f1.Underline = f2.Underline;
            //  Warning: assigning just one between the Subscript and the Superscript
            //  attributes, even to the msoFalse value, may impact the other. Therefore
            //  these attributes must be assigned only when strictly required.
            if (f2.Subscript == MsoTriState.msoTrue) {
                f1.Subscript = MsoTriState.msoTrue;
            }
        
            if (f2.Superscript == MsoTriState.msoTrue) {
                f1.Superscript = MsoTriState.msoTrue;
            }
        
            if ((!(f2.Subscript == MsoTriState.msoTrue)&& !(f2.Superscript == MsoTriState.msoTrue))) {
                f1.Subscript =  MsoTriState.msoFalse;
                f1.Superscript = MsoTriState.msoFalse;
            }
            this.assignColor(f1.Color, f2.Color);
        }

        //  This subroutine assigns the color in the ColorFormat object
        //  col2 to the ColorFormat object col1.
        //  Care must be taken in that the color may be specified as an
        //  index referring to the slide color scheme or as an RGB value.
        // 
        private void assignColor(PowerPoint.ColorFormat col1, PowerPoint.ColorFormat col2)
        {
            if (col2.Type != MsoColorType.msoColorTypeRGB)
            {
                //  I must protect from invalid assignments of color
                //  scheme indexes.
                try
                {
                    col1.SchemeColor = col2.SchemeColor;
                }
                catch (Exception)
                {
                    Trace.WriteLine("Assign Color :::::");
                }
            }
            else
            {
                col1.RGB = col2.RGB;
            }
        }
        private int findShape(Slide s, string id)
        {
            int i = 1;
            int findShape = 0;
            foreach (PowerPoint.Shape currentShape in s.Shapes)
            {
                if (currentShape.AlternativeText == id)
                {
                    findShape = i;
                    break;
                }
                i++;
            }
            return findShape;
        }
        public int ReturnAsciiCode(string value)
        {
            if (value.Length == 0)
                return 0;
            else if (value.Length > 1)
                value = value[0].ToString();
            int AsciiCodeO = (int)Convert.ToChar(value);
            return AsciiCodeO;
        }
    }
}
