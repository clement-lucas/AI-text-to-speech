using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace PptNarrator;

/// <summary>
/// Embeds audio (WAV) or video (MP4) files into PPTX slides with auto-play timing.
/// Audio is embedded as narration (hidden icon, plays automatically during slideshow).
/// Video is embedded as a visible element on the slide.
/// </summary>
static class SlideMediaEmbedder
{
    // 1×1 transparent PNG used as placeholder icon for audio / video poster
    private static readonly byte[] TransparentPixel = Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVQI12NgAAIABQAB" +
        "Nl7BcQAAAABJRU5ErkJggg==");

    /// <summary>
    /// Opens the PPTX at <paramref name="pptxPath"/> for editing and embeds
    /// media files into the specified slides.
    /// </summary>
    /// <param name="pptxPath">Path to the (copied) PPTX to modify.</param>
    /// <param name="slideMedia">Map of 0-based slide index → media file path.</param>
    /// <param name="mode">"audio" or "avatar".</param>
    public static void Embed(string pptxPath, Dictionary<int, string> slideMedia, string mode)
    {
        using var doc = PresentationDocument.Open(pptxPath, isEditable: true);
        var presentationPart = doc.PresentationPart
            ?? throw new InvalidOperationException("Invalid PPTX.");

        var slideIds = presentationPart.Presentation.SlideIdList?
            .ChildElements.OfType<SlideId>().ToList()
            ?? throw new InvalidOperationException("No slides.");

        foreach (var (slideIndex, mediaPath) in slideMedia)
        {
            if (slideIndex >= slideIds.Count) continue;

            var relId = slideIds[slideIndex].RelationshipId?.Value;
            if (relId is null) continue;

            var slidePart = (SlidePart)presentationPart.GetPartById(relId);

            if (mode == "avatar")
                EmbedVideo(doc, slidePart, mediaPath);
            else
                EmbedAudio(doc, slidePart, mediaPath);
        }

        doc.Save();
    }

    // ═══════════════════════════════════════════════════════════════════
    //  AUDIO NARRATION
    // ═══════════════════════════════════════════════════════════════════

    private static void EmbedAudio(PresentationDocument doc, SlidePart slidePart, string wavPath)
    {
        // 1. Add media data part
        var mediaPart = doc.CreateMediaDataPart("audio/wav", ".wav");
        using (var fs = File.OpenRead(wavPath))
            mediaPart.FeedData(fs);

        // 2. Create relationships from slide → media
        string audioRelId = slidePart.AddAudioReferenceRelationship(mediaPart).Id;
        string mediaRelId = slidePart.AddMediaReferenceRelationship(mediaPart).Id;

        // 3. Add placeholder icon image
        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(TransparentPixel))
            imagePart.FeedData(ms);
        string imageRelId = slidePart.GetIdOfPart(imagePart);

        // 4. Get unique shape ID
        uint shapeId = GetNextShapeId(slidePart);

        // 5. Add audio picture shape to the shape tree
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree.");

        shapeTree.Append(CreateAudioPicture(shapeId, audioRelId, mediaRelId, imageRelId));

        // 6. Set up auto-play narration timing
        slidePart.Slide.Timing = CreateNarrationTiming(shapeId);

        slidePart.Slide.Save();
    }

    private static Picture CreateAudioPicture(uint shapeId, string audioRelId, string mediaRelId, string imageRelId)
    {
        // Build NonVisualDrawingProperties with hlinkClick
        var cNvPr = new NonVisualDrawingProperties { Id = shapeId, Name = $"Narration {shapeId}" };
        cNvPr.AppendChild(new A.HyperlinkOnClick { Id = "", Action = "ppaction://media" });

        // NonVisualPictureDrawingProperties
        var cNvPicPr = new NonVisualPictureDrawingProperties();
        cNvPicPr.AppendChild(new A.PictureLocks { NoChangeAspect = true });

        // ApplicationNonVisualDrawingProperties with audioFile and p14:media extension
        var nvPr = new ApplicationNonVisualDrawingProperties();
        nvPr.AppendChild(new A.AudioFromFile { Link = audioRelId });
        nvPr.AppendChild(CreateMediaExtension(mediaRelId));

        var nvPicPr = new NonVisualPictureProperties();
        nvPicPr.Append(cNvPr, cNvPicPr, nvPr);

        // BlipFill with the placeholder icon
        var blipFill = new BlipFill(
            new A.Blip { Embed = imageRelId },
            new A.Stretch(new A.FillRectangle()));

        // ShapeProperties — small icon, positioned off-screen
        var spPr = new ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = 0, Y = 0 },
                new A.Extents { Cx = 304800, Cy = 304800 }), // 1 inch square
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

        return new Picture(nvPicPr, blipFill, spPr);
    }

    private static Timing CreateNarrationTiming(uint shapeId)
    {
        // "WithEffect" timing — audio auto-plays when the slide is shown (no click needed).
        return new Timing(
            new TimeNodeList(
                new ParallelTimeNode(
                    new CommonTimeNode(
                        new ChildTimeNodeList(
                            // Empty main sequence (required by PowerPoint)
                            new SequenceTimeNode(
                                new CommonTimeNode() { Id = 3, Duration = "indefinite", NodeType = TimeNodeValues.MainSequence },
                                new PreviousConditionList(
                                    new Condition(new TargetElement(new SlideTarget()))
                                    { Event = TriggerEventValues.OnPrevious, Delay = "0" }),
                                new NextConditionList(
                                    new Condition(new TargetElement(new SlideTarget()))
                                    { Event = TriggerEventValues.OnNext, Delay = "0" })
                            ) { Concurrent = true, NextAction = NextActionValues.Seek },
                            // WithEffect group — auto-plays on slide entry
                            new ParallelTimeNode(
                                new CommonTimeNode(
                                    new StartConditionList(
                                        new Condition { Delay = "0", Event = TriggerEventValues.OnBegin }
                                    ),
                                    new ChildTimeNodeList(
                                        new ParallelTimeNode(
                                            new CommonTimeNode(
                                                new StartConditionList(new Condition { Delay = "0" }),
                                                new ChildTimeNodeList(
                                                    new ParallelTimeNode(
                                                        new CommonTimeNode(
                                                            new StartConditionList(new Condition { Delay = "0" }),
                                                            new ChildTimeNodeList(
                                                                new Audio(
                                                                    new CommonMediaNode(
                                                                        new CommonTimeNode(
                                                                            new StartConditionList(new Condition { Delay = "0" })
                                                                        )
                                                                        { Id = 8, Fill = TimeNodeFillValues.Hold },
                                                                        new TargetElement(
                                                                            new ShapeTarget { ShapeId = shapeId.ToString() })
                                                                    ) { Volume = 80000 }
                                                                ) { IsNarration = true }
                                                            )
                                                        ) { Id = 7, Fill = TimeNodeFillValues.Hold }
                                                    )
                                                )
                                            ) { Id = 6, Fill = TimeNodeFillValues.Hold }
                                        )
                                    )
                                ) { Id = 5, Fill = TimeNodeFillValues.Hold, NodeType = TimeNodeValues.WithEffect }
                            )
                        )
                    ) { Id = 1, Duration = "indefinite", Restart = TimeNodeRestartValues.Never, NodeType = TimeNodeValues.TmingRoot }
                )
            )
        );
    }

    // ═══════════════════════════════════════════════════════════════════
    //  AVATAR VIDEO
    // ═══════════════════════════════════════════════════════════════════

    private static void EmbedVideo(PresentationDocument doc, SlidePart slidePart, string mp4Path)
    {
        // 1. Add media data part
        var mediaPart = doc.CreateMediaDataPart("video/mp4", ".mp4");
        using (var fs = File.OpenRead(mp4Path))
            mediaPart.FeedData(fs);

        // 2. Create relationships from slide → media
        string videoRelId = slidePart.AddVideoReferenceRelationship(mediaPart).Id;
        string mediaRelId = slidePart.AddMediaReferenceRelationship(mediaPart).Id;

        // 3. Add placeholder poster image
        var imagePart = slidePart.AddImagePart(ImagePartType.Png);
        using (var ms = new MemoryStream(TransparentPixel))
            imagePart.FeedData(ms);
        string imageRelId = slidePart.GetIdOfPart(imagePart);

        // 4. Get unique shape ID
        uint shapeId = GetNextShapeId(slidePart);

        // 5. Add video picture shape — bottom-right corner, 480×270 (16:9 at ~5in)
        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree
            ?? throw new InvalidOperationException("Slide has no shape tree.");

        // Format Video settings (from PowerPoint):
        //   Full video: 2.14×1.29 in
        //   Visible after crop: 1.13×1.29 in, Offset X=0, Y=0
        //   Crop position: Left=12.22 in, Top=0
        // Shape extents = visible (cropped) size on the slide.
        // fillRect right = -((fullWidth - visibleWidth) / visibleWidth) * 100000
        //                = -((2.14 - 1.13) / 1.13) * 100000 ≈ -89381
        // The negative value tells PowerPoint the image extends beyond the
        // shape’s right edge, giving Offset X = 0 (no left shift).
        long vidCx = 1033272;    // 1.13 inches (visible width)
        long vidCy = 1179576;    // 1.29 inches (visible height)
        long vidX  = 11173968;   // 12.22 inches from left
        long vidY  = 0;          // top = 0

        shapeTree.Append(CreateVideoPicture(shapeId, videoRelId, mediaRelId, imageRelId,
            vidX, vidY, vidCx, vidCy));

        // 6. Set up auto-play timing
        slidePart.Slide.Timing = CreateVideoTiming(shapeId);

        slidePart.Slide.Save();
    }

    private static Picture CreateVideoPicture(uint shapeId, string videoRelId, string mediaRelId,
        string imageRelId, long x, long y, long cx, long cy)
    {
        var cNvPr = new NonVisualDrawingProperties { Id = shapeId, Name = $"Avatar {shapeId}" };
        cNvPr.AppendChild(new A.HyperlinkOnClick { Id = "", Action = "ppaction://media" });

        var cNvPicPr = new NonVisualPictureDrawingProperties();
        cNvPicPr.AppendChild(new A.PictureLocks { NoChangeAspect = true });

        var nvPr = new ApplicationNonVisualDrawingProperties();
        nvPr.AppendChild(new A.VideoFromFile { Link = videoRelId });
        nvPr.AppendChild(CreateMediaExtension(mediaRelId));

        var nvPicPr = new NonVisualPictureProperties();
        nvPicPr.Append(cNvPr, cNvPicPr, nvPr);

        // BlipFill with fillRect — negative Right extends the image beyond
        // the shape’s right edge, cropping the right side of the video.
        // Offset X=0: left edge of image is flush with left edge of shape.
        var fillRect = new A.FillRectangle();
        fillRect.SetAttribute(new OpenXmlAttribute("r", "", "-89381"));

        var blipFill = new BlipFill(
            new A.Blip { Embed = imageRelId },
            new A.Stretch(fillRect));

        var spPr = new ShapeProperties(
            new A.Transform2D(
                new A.Offset { X = x, Y = y },
                new A.Extents { Cx = cx, Cy = cy }),
            new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle });

        return new Picture(nvPicPr, blipFill, spPr);
    }

    private static Timing CreateVideoTiming(uint shapeId)
    {
        // "WithEffect" timing — video auto-plays when the slide is shown (no click needed).
        // Structure: tmRoot → childTnLst → seq(mainSeq) + par(withEffect)
        //   - The mainSeq is empty (no click animations)
        //   - The withEffect par contains the video that starts on slideBegin
        return new Timing(
            new TimeNodeList(
                new ParallelTimeNode(
                    new CommonTimeNode(
                        new ChildTimeNodeList(
                            // Empty main sequence (required by PowerPoint)
                            new SequenceTimeNode(
                                new CommonTimeNode() { Id = 3, Duration = "indefinite", NodeType = TimeNodeValues.MainSequence },
                                new PreviousConditionList(
                                    new Condition(new TargetElement(new SlideTarget()))
                                    { Event = TriggerEventValues.OnPrevious, Delay = "0" }),
                                new NextConditionList(
                                    new Condition(new TargetElement(new SlideTarget()))
                                    { Event = TriggerEventValues.OnNext, Delay = "0" })
                            ) { Concurrent = true, NextAction = NextActionValues.Seek },
                            // WithEffect group — auto-plays on slide entry
                            new ParallelTimeNode(
                                new CommonTimeNode(
                                    new StartConditionList(
                                        new Condition { Delay = "0", Event = TriggerEventValues.OnBegin }
                                    ),
                                    new ChildTimeNodeList(
                                        new ParallelTimeNode(
                                            new CommonTimeNode(
                                                new StartConditionList(new Condition { Delay = "0" }),
                                                new ChildTimeNodeList(
                                                    new ParallelTimeNode(
                                                        new CommonTimeNode(
                                                            new StartConditionList(new Condition { Delay = "0" }),
                                                            new ChildTimeNodeList(
                                                                new Video(
                                                                    new CommonMediaNode(
                                                                        new CommonTimeNode(
                                                                            new StartConditionList(new Condition { Delay = "0" })
                                                                        )
                                                                        { Id = 8, Fill = TimeNodeFillValues.Hold },
                                                                        new TargetElement(
                                                                            new ShapeTarget { ShapeId = shapeId.ToString() })
                                                                    ) { Volume = 80000 }
                                                                ) { FullScreen = false }
                                                            )
                                                        ) { Id = 7, Fill = TimeNodeFillValues.Hold }
                                                    )
                                                )
                                            ) { Id = 6, Fill = TimeNodeFillValues.Hold }
                                        )
                                    )
                                ) { Id = 5, Fill = TimeNodeFillValues.Hold, NodeType = TimeNodeValues.WithEffect }
                            )
                        )
                    ) { Id = 1, Duration = "indefinite", Restart = TimeNodeRestartValues.Never, NodeType = TimeNodeValues.TmingRoot }
                )
            )
        );
    }

    // ═══════════════════════════════════════════════════════════════════
    //  HELPERS
    // ═══════════════════════════════════════════════════════════════════

    /// <summary>
    /// Creates the p14:media extension element for PowerPoint 2010+ compatibility.
    /// </summary>
    private static ApplicationNonVisualDrawingPropertiesExtensionList CreateMediaExtension(string mediaRelId)
    {
        // Build: <p:extLst><p:ext uri="{DAA...}"><p14:media r:embed="rId"/></p:ext></p:extLst>
        var mediaElement = new P14.Media();
        mediaElement.SetAttribute(new OpenXmlAttribute("r", "embed",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships", mediaRelId));

        var ext = new Extension { Uri = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}" };
        ext.AppendChild(mediaElement);

        var extList = new ApplicationNonVisualDrawingPropertiesExtensionList();
        extList.AppendChild(ext);
        return extList;
    }

    /// <summary>
    /// Finds the highest shape ID in the slide and returns the next available one.
    /// </summary>
    private static uint GetNextShapeId(SlidePart slidePart)
    {
        uint maxId = 0;

        var shapeTree = slidePart.Slide.CommonSlideData?.ShapeTree;
        if (shapeTree is null) return 100;

        foreach (var elem in shapeTree.Descendants())
        {
            if (elem is NonVisualDrawingProperties nvPr && nvPr.Id is not null && nvPr.Id.Value > maxId)
                maxId = nvPr.Id.Value;
        }

        // Also check shapes referenced in notes part, etc. Just ensure it's high enough.
        return Math.Max(maxId + 1, 100);
    }
}
