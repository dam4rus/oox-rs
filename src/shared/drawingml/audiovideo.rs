use crate::{
    error::{MissingAttributeError, MissingChildNodeError, NotGroupMemberError},
    shared::relationship::RelationshipId,
    xml::XmlNode,
    xsdtypes::{XsdChoice, XsdType},
};

pub type Result<T> = ::std::result::Result<T, Box<dyn (::std::error::Error)>>;

#[derive(Debug, Clone, PartialEq)]
pub struct AudioCD {
    /// This element specifies the start point for a CD Audio sound element. Encompassed within this element are the
    /// time and track at which the sound should begin its playback. This element is used in conjunction with an Audio
    /// End Time element to specify the time span for an entire audioCD sound element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:audioCd>
    ///   <a:st track="1" time="2"/>
    ///   <a:end track="3" time="65"/>
    /// </a:audioCd>
    /// ```
    ///
    /// In the above example, the audioCD sound element shown specifies for a portion of audio spanning from 2
    /// seconds into the first track to 1 minute, 5 seconds into the third track.
    pub start_time: AudioCDTime,

    /// This element specifies the end point for a CD Audio sound element. Encompassed within this element are the
    /// time and track at which the sound should halt its playback. This element is used in conjunction with an Audio
    /// Start Time element to specify the time span for an entire audioCD sound element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <a:audioCd>
    ///   <a:st track="1" time="2"/>
    ///   <a:end track="3" time="65"/>
    /// </a:audioCd>
    /// ```
    ///
    /// In the above example, the audioCD sound element shown specifies for a portion of audio spanning from 2
    /// seconds into the first track to 1 minute, 5 seconds into the third track.
    pub end_time: AudioCDTime,
}

impl AudioCD {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut start_time = None;
        let mut end_time = None;

        for child_node in &xml_node.child_nodes {
            match child_node.local_name() {
                "st" => start_time = Some(AudioCDTime::from_xml_element(child_node)?),
                "end" => end_time = Some(AudioCDTime::from_xml_element(child_node)?),
                _ => (),
            }
        }

        let start_time = start_time.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "st"))?;
        let end_time = end_time.ok_or_else(|| MissingChildNodeError::new(xml_node.name.clone(), "end"))?;

        Ok(Self { start_time, end_time })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct AudioFile {
    /// Specifies the identification information for a linked object. This attribute is used to
    /// specify the location of an object that does not reside within this file.
    pub link: RelationshipId,

    /// Specifies the content type for the external file that is referenced by this element. Content
    /// types define a media type, a subtype, and an optional set of parameters, as defined in
    /// Part 2. If a rendering application cannot process external content of the content type
    /// specified, then the specified content can be ignored.
    ///
    /// If this attribute is omitted, application should attempt to determine the content type by
    /// reading the contents of the relationship’s target.
    ///
    /// Suggested audio types:
    /// * aiff
    /// * midi
    /// * ogg
    /// * mpeg
    ///
    /// A producer that wants interoperability should use the following standard format:
    /// * audio
    /// * mpeg ISO
    /// * IEC 11172-3
    pub content_type: Option<String>,
}

impl AudioFile {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut link = None;
        let mut content_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r:link" => link = Some(value.clone()),
                "contentType" => content_type = Some(value.clone()),
                _ => (),
            }
        }

        let link = link.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:link"))?;

        Ok(Self { link, content_type })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct AudioCDTime {
    /// Specifies which track of the CD this Audio begins playing on. This attribute is required and
    /// cannot be omitted.
    pub track: u8,

    /// Specifies the time in seconds that the CD Audio should be started at.
    ///
    /// Defaults to 0
    pub time: Option<u32>,
}

impl AudioCDTime {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut track = None;
        let mut time = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "track" => track = Some(value.parse()?),
                "time" => time = Some(value.parse()?),
                _ => (),
            }
        }

        let track = track.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "track"))?;

        Ok(Self { track, time })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct QuickTimeFile {
    /// Specifies the identification information for a linked object. This attribute is used to
    /// specify the location of an object that does not reside within this file.
    pub link: RelationshipId,
}

impl QuickTimeFile {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let link = xml_node
            .attributes
            .get("r:link")
            .ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:link"))?
            .clone();

        Ok(Self { link })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct VideoFile {
    /// Specifies the identification information for a linked video file. This attribute is used to
    /// specify the location of an object that does not reside within this file.
    pub link: RelationshipId,

    /// Specifies the content type for the external file that is referenced by this element. Content
    /// types define a media type, a subtype, and an optional set of parameters, as defined in
    /// Part 2. If a rendering application cannot process external content of the content type
    /// specified, then the specified content can be ignored.
    ///
    /// Suggested video formats:
    /// * avi
    /// * mpg
    /// * mpeg
    /// * ogg
    /// * quicktime
    /// * vc1
    ///
    /// If this attribute is omitted, application should attempt to determine the content type by
    /// reading the contents of the relationship’s target.
    pub content_type: Option<String>,
}

impl VideoFile {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut link = None;
        let mut content_type = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r:link" => link = Some(value.clone()),
                "contentType" => content_type = Some(value.clone()),
                _ => (),
            }
        }

        let link = link.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:link"))?;

        Ok(Self { link, content_type })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub struct EmbeddedWAVAudioFile {
    /// Specifies the identification information for an embedded audio file. This attribute is used
    /// to specify the location of an object that resides locally within the file.
    pub embed_rel_id: RelationshipId,

    /// Specifies the original name or given short name for the corresponding sound. This is used
    /// to distinguish this sound from others by providing a human readable name for the
    /// attached sound should the user need to identify the sound among others within the UI.
    pub name: Option<String>,
    //pub built_in: Option<bool>, // false
}

impl EmbeddedWAVAudioFile {
    pub fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        let mut embed_rel_id = None;
        let mut name = None;

        for (attr, value) in &xml_node.attributes {
            match attr.as_str() {
                "r:embed" => embed_rel_id = Some(value.clone()),
                "name" => name = Some(value.clone()),
                _ => (),
            }
        }

        let embed_rel_id = embed_rel_id.ok_or_else(|| MissingAttributeError::new(xml_node.name.clone(), "r:embed"))?;

        Ok(Self { embed_rel_id, name })
    }
}

#[derive(Debug, Clone, PartialEq)]
pub enum Media {
    /// This element specifies the existence of Audio from a CD. This element is specified within the non-visual
    /// properties of an object. The audio shall be attached to an object as this is how it is represented within the
    /// document. The actual playing of the sound however is done within the timing node list that is specified under
    /// the timing element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="7" name="Rectangle 6">
    ///       <a:hlinkClick r:id="" action="ppaction://media"/>
    ///     </p:cNvPr>
    ///     <p:cNvPicPr>
    ///       <a:picLocks noRot="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr>
    ///       <a:audioCd>
    ///         <a:st track="1"/>
    ///         <a:end track="3" time="65"/>
    ///       </a:audioCd>
    ///     </p:nvPr>
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    ///
    /// In the above example, we see that there is a single audioCD element attached to this picture. This picture is
    /// placed within the document just as a normal picture or shape would be. The id of this picture, namely 7 in this
    /// case, is used to refer to this audioCD element from within the timing node list. For this example we see that the
    /// audio for this CD starts playing at the 0 second mark on the first track and ends on the 1 minute 5 second mark
    /// of the third track.
    AudioCd(AudioCD),

    /// This element specifies the existence of an audio WAV file. This element is specified within the non-visual
    /// properties of an object. The audio shall be attached to an object as this is how it is represented within the
    /// document. The actual playing of the audio however is done within the timing node list that is specified under the
    /// timing element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="7" name="Rectangle 6">
    ///       <a:hlinkClick r:id="" action="ppaction://media"/>
    ///     </p:cNvPr>
    ///     <p:cNvPicPr>
    ///       <a:picLocks noRot="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr>
    ///       <a:wavAudioFile r:embed="rId2"/>
    ///     </p:nvPr>
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    ///
    /// In the above example, we see that there is a single wavAudioFile element attached to this picture. This picture
    /// is placed within the document just as a normal picture or shape would be. The id of this picture, namely 7 in this
    /// case, is used to refer to this wavAudioFile element from within the timing node list. The Embedded relationship
    /// id is used to retrieve the actual audio file for playback purposes.
    WavAudioFile(EmbeddedWAVAudioFile),

    /// This element specifies the existence of an audio file. This element is specified within the non-visual properties of
    /// an object. The audio shall be attached to an object as this is how it is represented within the document. The
    /// actual playing of the audio however is done within the timing node list that is specified under the timing
    /// element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="7" name="Rectangle 6">
    ///       <a:hlinkClick r:id="" action="ppaction://media"/>
    ///     </p:cNvPr>
    ///     <p:cNvPicPr>
    ///       <a:picLocks noRot="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr>
    ///       <a:audioFile r:link="rId1"/>
    ///     </p:nvPr>
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    ///
    /// In the above example, we see that there is a single audioFile element attached to this picture. This picture is
    /// placed within the document just as a normal picture or shape would be. The id of this picture, namely 7 in this
    /// case, is used to refer to this audioFile element from within the timing node list. The Linked relationship id is
    /// used to retrieve the actual audio file for playback purposes.
    AudioFile(AudioFile),

    /// This element specifies the existence of a video file. This element is specified within the non-visual properties of
    /// an object. The video shall be attached to an object as this is how it is represented within the document. The
    /// actual playing of the video however is done within the timing node list that is specified under the timing
    /// element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="7" name="Rectangle 6">
    ///       <a:hlinkClick r:id="" action="ppaction://media"/>
    ///     </p:cNvPr>
    ///     <p:cNvPicPr>
    ///       <a:picLocks noRot="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr>
    ///       <a:videoFile r:link="rId1"/>
    ///     </p:nvPr>
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    ///
    /// In the above example, we see that there is a single videoFile element attached to this picture. This picture is
    /// placed within the document just as a normal picture or shape would be. The id of this picture, namely 7 in this
    /// case, is used to refer to this videoFile element from within the timing node list. The Linked relationship id is
    /// used to retrieve the actual video file for playback purposes.
    VideoFile(VideoFile),

    /// This element specifies the existence of a QuickTime file. This element is specified within the non-visual
    /// properties of an object. The QuickTime file shall be attached to an object as this is how it is represented
    /// within the document. The actual playing of the QuickTime however is done within the timing node list that is
    /// specified under the timing element.
    ///
    /// # Xml example
    ///
    /// ```xml
    /// <p:pic>
    ///   <p:nvPicPr>
    ///     <p:cNvPr id="7" name="Rectangle 6">
    ///       <a:hlinkClick r:id="" action="ppaction://media"/>
    ///     </p:cNvPr>
    ///     <p:cNvPicPr>
    ///       <a:picLocks noRot="1"/>
    ///     </p:cNvPicPr>
    ///     <p:nvPr>
    ///       <a:quickTimeFile r:link="rId1"/>
    ///     </p:nvPr>
    ///   </p:nvPicPr>
    ///   ...
    /// </p:pic>
    /// ```
    ///
    /// In the above example, we see that there is a single quickTimeFile element attached to this picture. This picture
    /// is placed within the document just as a normal picture or shape would be. The id of this picture, namely 7 in this
    /// case, is used to refer to this quickTimeFile element from within the timing node list. The Linked relationship id
    /// is used to retrieve the actual video file for playback purposes.
    QuickTimeFile(QuickTimeFile),
}

impl XsdType for Media {
    fn from_xml_element(xml_node: &XmlNode) -> Result<Self> {
        match xml_node.local_name() {
            "audioCd" => Ok(Media::AudioCd(AudioCD::from_xml_element(xml_node)?)),
            "wavAudioFile" => Ok(Media::WavAudioFile(EmbeddedWAVAudioFile::from_xml_element(xml_node)?)),
            "audioFile" => Ok(Media::AudioFile(AudioFile::from_xml_element(xml_node)?)),
            "videoFile" => Ok(Media::VideoFile(VideoFile::from_xml_element(xml_node)?)),
            "quickTimeFile" => Ok(Media::QuickTimeFile(QuickTimeFile::from_xml_element(xml_node)?)),
            _ => Err(Box::new(NotGroupMemberError::new(xml_node.name.clone(), "EG_Media"))),
        }
    }
}

impl XsdChoice for Media {
    fn is_choice_member<T>(name: T) -> bool
    where
        T: AsRef<str>,
    {
        match name.as_ref() {
            "audioCd" | "wavAudioFile" | "audioFile" | "videoFile" | "quickTimeFile" => true,
            _ => false,
        }
    }
}
