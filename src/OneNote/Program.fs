
open System
open System.Diagnostics
open System.Drawing
open System.Globalization
open System.IO
open System.Net
open System.Windows.Forms
open System.Text
open System.Text.RegularExpressions
open System.Threading
open System.Xml

[<AutoOpen>]
module AutoOpen =
  let inline encode (s : string) = WebUtility.UrlEncode s
  let inline sz x y = Size (x, y)
  let culture = CultureInfo.InvariantCulture

module Authenticate =
  let scope         = encode  (
                                String.concat 
                                  " " 
                                  [|
                                    "wl.signin" 
                                    "wl.basic"
                                    "office.onenote_create"
                                    "office.onenote"
                                    "office.onenote_update"
                                    "office.onenote_update_by_app"
                                  |]
                              )
  let responseType  = encode "token"
  let redirectUri   = encode "https://login.live.com/oauth20_desktop.srf"
  let logonUrl clientId = 
    sprintf 
      "https://login.live.com/oauth20_authorize.srf?client_id=%s&scope=%s&response_type=%s&redirect_uri=%s"
      (encode clientId)
      scope
      responseType
      redirectUri

  let keyValueRegex = 
    Regex (
        "^(?<key>[^=]*)\=(?<value>.*)$"
      , RegexOptions.Compiled ||| RegexOptions.CultureInvariant ||| RegexOptions.Singleline ||| RegexOptions.ExplicitCapture
      )

  let lift (f : 'T -> 'U option) (o : 'T option) : 'U option =
    match o with
    | Some v -> f v
    | _ -> None

  let parseInteger (s : string) : int option =
    let r,v = Int32.TryParse (s, NumberStyles.Integer, CultureInfo.InvariantCulture)
    if r then Some v
    else None

  let parseKeyValue (kv : string) : (string*string) option =
    let m = keyValueRegex.Match kv
    if m.Success then
      Some (m.Groups.["key"].Value, m.Groups.["value"].Value)
    else
      None

  let parseFragment (fragment : string) : (string*string) [] option =
    if fragment.Length > 0 && fragment.[0] = '#' then
      let kvs = 
        fragment.Substring(1).Split '&'
        |> Seq.choose parseKeyValue
        |> Seq.toArray

      Some kvs
    else
      None

  let findValue (k : string) (kvs : (string*string) []) : string option =
    kvs 
    |> Array.tryFind (fun (kk,vv) -> k = kk)
    |> Option.map (fun (kk,vv) -> vv)

  let showForm (form : Form) =
    Async.FromContinuations <| fun (callback, e, c) ->
      // Close form on exception/cancellation
      let subscription = ref (None : IDisposable option)
      let onClosed cc = 
        match !subscription with
        | Some d -> d.Dispose ()
        | _ -> ()
        callback ()
      try
        subscription := Some <| form.Closed.Subscribe onClosed
        form.Show ()
      with
      | ex -> e ex

  let logon clientId = 
    async {
      let result = ref None

      use container   = new ComponentModel.Container ()
      use form        = new Form (  Text        = "Authenticate to Windows Live"
                                  , Size        = sz 800 600
                                  , MinimumSize = sz 800 600
                                  )
      use webBrowser  = new WebBrowser (  Dock  = DockStyle.Fill
                                        , Url   = Uri (logonUrl clientId)
                                        )

      use closeTimer = new Windows.Forms.Timer (container, Interval = 1000)

      use navigatedSubscriber = webBrowser.Navigated.Subscribe (fun e -> 
        let url     = e.Url
        let schema  = url.Scheme
        let host    = url.Host
        let path    = url.AbsolutePath
        let fragment= url.Fragment
        let kvs     = parseFragment fragment
        let f k     = kvs |> lift (findValue k)
       
        let accessToken         = f "access_token"
        let tokenType           = f "token_type"
        let authenticationToken = f "authentication_token"
        let expiresIn           = f "expires_in" |> lift parseInteger
        let userId              = f "user_id"
        let scope               = f "scope"

        match schema, host, path, accessToken, authenticationToken with
        | "https", "login.live.com", "/oauth20_desktop.srf", Some accessToken, Some authenticationToken ->
          result := Some (accessToken, authenticationToken)
          closeTimer.Enabled <- true
        | _ -> ()
        )

      use tickSubscriber = closeTimer.Tick.Subscribe (fun e -> closeTimer.Enabled <- false; form.Close ())

      form.Controls.Add webBrowser

      do! showForm form

      return
        match !result with
        | Some tokens -> tokens
        | _ -> failwith "No accesstoken set"
    }

module ApiEndPoint =

  [<Literal>]
  let uriNotebooks      = "https://www.onenote.com/api/v1.0/notebooks"
  [<Literal>]
  let uriPages          = "https://www.onenote.com/api/v1.0/pages"
  [<Literal>]
  let uriSections       = "https://www.onenote.com/api/v1.0/sections"
  [<Literal>]
  let uriSectionGroups  = "https://www.onenote.com/api/v1.0/sectiongroups"

  type NotebookId     = NotebookId      of string
  type PageId         = PageId          of string
  type SectionId      = SectionId       of string
  type SectionGroupId = SectionGroupId  of string

  type GetEndPoint =
    | Notebooks
    | Notebook                    of NotebookId
    | Notebook_Sections           of NotebookId
    | Notebook_SectionGroups      of NotebookId
    | Pages   
    | Page                        of PageId
    | Page_Content                of PageId*bool
    | Sections
    | Section                     of SectionId
    | Section_Pages               of SectionId
    | SectionGroups
    | SectionGroup                of SectionGroupId
    | SectionGroup_SectionGroups  of SectionGroupId
    | SectionGroup_Sections       of SectionGroupId

  let getEndPointUri = function
    | Notebooks                                       -> uriNotebooks
    | Notebook (NotebookId id)                        -> sprintf  "%s/%s" uriNotebooks (encode id)
    | Notebook_Sections (NotebookId id)               -> sprintf  "%s/%s/sections" uriNotebooks (encode id)
    | Notebook_SectionGroups (NotebookId id)          -> sprintf  "%s/%s/sectiongroups" uriNotebooks (encode id)
    | Pages                                           -> uriPages
    | Page (PageId id)                                -> sprintf  "%s/%s" uriPages (encode id)
    | Page_Content (PageId id, includeIDs)            -> sprintf  "%s/%s/content?includeIDs=%s" uriPages (encode id) (if includeIDs then "true" else "false")
    | Sections                                        -> uriSections
    | Section (SectionId id)                          -> sprintf  "%s/%s" uriSections (encode id)
    | Section_Pages (SectionId id)                    -> sprintf  "%s/%s/pages" uriSections (encode id)
    | SectionGroups                                   -> uriSectionGroups
    | SectionGroup (SectionGroupId id)                -> sprintf  "%s/%s" uriSectionGroups (encode id)
    | SectionGroup_SectionGroups (SectionGroupId id)  -> sprintf  "%s/%s/sectionGroups" uriSectionGroups (encode id)
    | SectionGroup_Sections (SectionGroupId id)       -> sprintf  "%s/%s/sections" uriSectionGroups (encode id)

  let getRequest accessToken getEndPoint =
    async {
      let uri = Uri (getEndPointUri getEndPoint)
      let webRequest : HttpWebRequest = downcast HttpWebRequest.Create uri
      webRequest.Headers.Add ("Authorization", sprintf "Bearer %s" accessToken)

      use! webResponse = webRequest.AsyncGetResponse ()
      use responseStream = webResponse.GetResponseStream ()
      use ss = new StreamReader (responseStream)
      let! content = Async.AwaitTask (ss.ReadToEndAsync ())

      return content
    }

  let getNotebooks                  accessToken               = getRequest accessToken Notebooks
  let getNotebook                   accessToken id            = getRequest accessToken (Notebook id)
  let getNotebook_SectionGroups     accessToken id            = getRequest accessToken (Notebook_SectionGroups id)
  let getNotebook_Sections          accessToken id            = getRequest accessToken (Notebook_Sections id)

  let getPages                      accessToken               = getRequest accessToken Pages
  let getPage                       accessToken id            = getRequest accessToken (Page id)
  let getPage_Content               accessToken id includeIDs = getRequest accessToken (Page_Content (id, includeIDs))

  let getSections                   accessToken               = getRequest accessToken Sections
  let getSection                    accessToken id            = getRequest accessToken (Section id)
  let getSection_Pages              accessToken id            = getRequest accessToken (Section_Pages id)

  let getSectionGroups              accessToken               = getRequest accessToken SectionGroups
  let getSectionGroup               accessToken id            = getRequest accessToken (SectionGroup id)
  let getSectionGroup_SectionGroups accessToken id            = getRequest accessToken (SectionGroup_SectionGroups id)
  let getSectionGroup_Sections      accessToken id            = getRequest accessToken (SectionGroup_SectionGroups id)

module PageModel =

  type Color = Color of string

  type FontFamily = FontFamily of string

  type FontSize =
    | FontSizePx of decimal
    | FontSizePt of decimal

  type FontStyle =
    | Normal
    | Italic

  type FontWeight =
    | Normal
    | Bold

  type TextAlign =
    | Left
    | Center
    | Right

  type TextDecoration =
    | None
    | LineThrough
    | Underline

  type StyleSet =
    {
      BackgroundColor : Color option
      Color           : Color option
      FontFamily      : FontFamily option
      FontSize        : FontSize option
      FontStyle       : FontStyle option
      FontWeight      : FontWeight option
      MaxWidth        : int option
      MaxHeight       : int option
      TextAlign       : TextAlign option
      TextDecoration  : TextDecoration option
      CustomStyle     : (string*string) option
    }

    static member Empty : StyleSet =
      {
        BackgroundColor = Option.None
        Color           = Option.None
        FontFamily      = Option.None
        FontSize        = Option.None
        FontStyle       = Option.None
        FontWeight      = Option.None
        MaxWidth        = Option.None
        MaxHeight       = Option.None
        TextAlign       = Option.None
        TextDecoration  = Option.None
        CustomStyle     = Option.None
      }
  type Styles = 
    | NoStyle
    | StyleSet of StyleSet

  type Heading =
    | H1
    | H2
    | H3
    | H4
    | H5
    | H6

  type ListKind =
    | Ordered
    | Unordered

  type CharacterStyle =
    | Bold
    | Emphasized
    | Strong
    | Italic
    | Underline
    | StrikeThrough
    | Deleted
    | Superscript
    | Subscript
    | Citation
    | Font          of FontFamily*Color

  type TableColumn = 
    | TableColumn of Styles*Tags
  and TableHeader = 
    | TableHeader of Styles*Tags
  and TableHeaders = Styles*TableHeader []
  and TableRow = 
    | TableRow of Styles*TableColumn []
  and TableRows = TableRow []
  and ListItem = ListItem of Styles*Tags
  and ListItems = ListItem []
  and TagId = 
    | NoId
    | TagId of string
  and TagAttribute = 
    | TagAttribute of string*string
  and TagAttributes = TagAttribute []
  and TagKind =
    | Div               
    | Span              
    | Anchor            of Uri*string
    | Paragraph         
    | Break
    | Heading           of Heading
    | List              of ListKind*ListItems
    | Image             of Uri*string*int*int
    | Object
    | Table             of TableHeaders*TableRows
    | CharacterStyle    of CharacterStyle
    | Text              of string
    | Unrecognized      of string
  and Tag =
    | Tag of TagKind*TagId*Styles*Tags
  and Tags = Tag []
  
  type Page =
    {
      Title     : string
      Created   : DateTime
      Body      : Styles*Tags
    }

    static member New title created styles tags : Page =
      {
        Title     = title
        Created   = created
        Body      = (styles, tags)
      }

  module Details =
    type TagBehavior =
      | Content
      | Leaf
      | Node

    let tagMeta (tagKind : TagKind) : TagBehavior*string = 
      match tagKind with
      | Div                         -> Node   , "div"
      | Span                        -> Node   , "span"
      | Anchor _                    -> Leaf   , "a"
      | Paragraph                   -> Node   , "p"
      | Break                       -> Leaf   , "br"
      | Heading H1                  -> Node   , "h1"
      | Heading H2                  -> Node   , "h2"
      | Heading H3                  -> Node   , "h3"
      | Heading H4                  -> Node   , "h4"
      | Heading H5                  -> Node   , "h5"
      | Heading H6                  -> Node   , "h6"
      | List (Ordered, _)           -> Leaf   , "ol"
      | List (Unordered, _)         -> Leaf   , "ul"
      | Image _                     -> Leaf   , "img"
      | Object                      -> Leaf   , "object"
      | Table _                     -> Leaf   , "table"
      | CharacterStyle Bold         -> Node   , "b"
      | CharacterStyle Emphasized   -> Node   , "em"
      | CharacterStyle Strong       -> Node   , "strong"
      | CharacterStyle Italic       -> Node   , "i"
      | CharacterStyle Underline    -> Node   , "u"
      | CharacterStyle StrikeThrough-> Node   , "strike"
      | CharacterStyle Deleted      -> Node   , "del"
      | CharacterStyle Superscript  -> Node   , "sup"
      | CharacterStyle Subscript    -> Node   , "sub"
      | CharacterStyle Citation     -> Node   , "cite"
      | CharacterStyle (Font _)     -> Node   , "font"
      | Text _                      -> Content, ""
      | Unrecognized tag            -> Node   , tag

    let renderStyles (xtw : XmlWriter) (styles : Styles) : unit =
      match styles with
      | NoStyle -> ()
      | StyleSet set ->
        let sb = StringBuilder ()

        let prepend = ref ""

        let inline append (s : string) =
          ignore <| sb.Append s

        let lift (ov : 'T option) (name : string) (appender : 'T -> unit) =
          match ov with
          | Some v -> 
            append (!prepend)
            append name
            append ":"
            prepend := ";"
            appender v

          | _ -> ()

        lift set.BackgroundColor  "bgcolor"         <| fun (Color s) -> append s
        lift set.Color            "color"           <| fun (Color s) -> append s
        lift set.FontFamily       "font-family"     <| fun (FontFamily s) -> append s
        lift set.FontSize         "font-size"       <| fun sz -> 
          match sz with
          | FontSizePt d -> 
            append (d.ToString culture)
            append "px"
          | FontSizePx d -> 
            append (d.ToString culture)
            append "pt"
        lift set.FontStyle        "font-style"      <| fun fs ->
          match fs with
          | FontStyle.Normal -> append "normal"
          | FontStyle.Italic -> append "italic"
        lift set.FontWeight       "font-weight"     <| fun fw -> 
          match fw with
          | FontWeight.Normal -> append "normal"
          | FontWeight.Bold -> append "bold"
        lift set.MaxWidth         "max-width"       <| fun w -> append (w.ToString culture)
        lift set.MaxHeight        "max-height"      <| fun h -> append (h.ToString culture)
        lift set.TextAlign        "text-align"      <| fun ta ->
          match ta with
          | Left -> append "left"
          | Center -> append "center"
          | Right -> append "right"
        lift set.TextDecoration   "text-decoration" <| fun td -> 
          match td with
          | TextDecoration.None -> append "none"
          | TextDecoration.LineThrough -> append "line-through"
          | TextDecoration.Underline -> append "underline"
// TODO:
//        lift set.CustomStyle      "bgcolor"     <| fun (Color s) -> append s

        xtw.WriteAttributeString ("style", sb.ToString ())

    let renderTagId (xtw : XmlWriter) (tagId : TagId) : unit =
      ()

    let rec renderTagKind (xtw : XmlWriter) (tagKind : TagKind) : unit =
      match tagKind with
      | Div                 -> ()
      | Span                -> ()
      | Anchor (href, title)->
        xtw.WriteAttributeString ("href", string href)
        xtw.WriteString title
      | Paragraph                       -> ()
      | Break                           -> ()
      | Heading _                       -> ()
      | List (_, listItems)              ->
        for (ListItem (styles, tags)) in listItems do
          xtw.WriteStartElement "li"
          renderStyles xtw styles
          renderTags xtw tags
          xtw.WriteEndElement ()
      | Image (src, alt, width, height) ->
        xtw.WriteAttributeString ("src"     , string src    )
        xtw.WriteAttributeString ("alt"     , alt           )
        xtw.WriteAttributeString ("width"   , string width  )
        xtw.WriteAttributeString ("height"  , string height )
      | Object                          -> ()
      | Table (th, trows)         ->
          let styles, theaders = th

          if theaders.Length > 0 then
            xtw.WriteStartElement "tr"
            renderStyles xtw styles
            for (TableHeader (styles,tags)) in theaders do
              xtw.WriteStartElement "th"
              renderStyles xtw styles
              renderTags xtw tags
              xtw.WriteEndElement ()
            xtw.WriteEndElement ()

          if trows.Length > 0 then
            for (TableRow (styles,trow)) in trows do
              xtw.WriteStartElement "tr"
              renderStyles xtw styles
              for (TableColumn (styles,tags)) in trow do
                xtw.WriteStartElement "td"
                renderStyles xtw styles
                renderTags xtw tags
                xtw.WriteEndElement ()
              xtw.WriteEndElement ()
      | CharacterStyle _                -> ()
      | Text s                          -> xtw.WriteString s
      | Unrecognized _                  -> ()

    and renderTag (xtw : XmlWriter) (Tag (tagKind, tagId, styles, tags)) : unit =

      let tagBehavior, tagName = tagMeta tagKind

      let hasElement = 
        match tagBehavior with
        | Content -> false
        | _ -> true

      if hasElement then
        xtw.WriteStartElement tagName

        renderTagId xtw tagId
        renderStyles xtw styles

      renderTagKind xtw tagKind

      match tagBehavior with
      | Content -> ()
      | Leaf -> ()  // TODO: Raise error on non-empty tags
      | Node -> renderTags xtw tags

      if hasElement then
        xtw.WriteEndElement ()

    and renderTags (xtw : XmlWriter) (tags : Tags) : unit =
      for tag in tags do
        renderTag xtw tag

    let noStyle                 = NoStyle
    let noTags                  = [||]

    let inline makeTag tk tags  = Tag (tk, NoId, noStyle, tags)

    let inline liftStyle (f : StyleSet -> StyleSet) (Tag (tagKind, id, styles, tags)) = 
      let ss : StyleSet =
        match styles with
        | NoStyle -> StyleSet.Empty
        | StyleSet ss -> ss
      let nss = f ss
        
      Tag (tagKind, id, StyleSet nss, tags)


  open Details

  let renderPage (page : Page) : string =
    let settings = XmlWriterSettings (Indent = true)
    let sb = new StringBuilder ()
    use sw = new StringWriter (sb, culture)
    use xtw = new XmlTextWriter (sw)
    
    xtw.WriteStartElement "html"

    xtw.WriteStartElement "head"
    xtw.WriteElementString ("title", page.Title)
    xtw.WriteElementString ("create", page.Created.ToString ("s", culture))
    xtw.WriteEndElement ()

    let styles, tags = page.Body

    xtw.WriteStartElement "body"
    renderStyles xtw styles
    renderTags xtw tags

    xtw.WriteEndElement ()

    xtw.WriteEndElement ()
    
    sw.ToString ()

  let page title styles tags    = Page.New title DateTime.Now styles tags

  let div tags                  = makeTag TagKind.Div                                 tags
  let span tags                 = makeTag TagKind.Span                                tags
  let a uri title               = makeTag (TagKind.Anchor (uri, title))               noTags
  let p tags                    = makeTag TagKind.Paragraph                           tags
  let br                        = makeTag TagKind.Break                               noTags
  let h1 tags                   = makeTag (TagKind.Heading H1)                        tags
  let h2 tags                   = makeTag (TagKind.Heading H2)                        tags
  let h3 tags                   = makeTag (TagKind.Heading H3)                        tags
  let h4 tags                   = makeTag (TagKind.Heading H4)                        tags
  let h5 tags                   = makeTag (TagKind.Heading H5)                        tags
  let h6 tags                   = makeTag (TagKind.Heading H6)                        tags
  let ol lis                    = makeTag (TagKind.List (Ordered, lis))               noTags
  let ul lis                    = makeTag (TagKind.List (Unordered, lis))             noTags
  let img uri alt width height  = makeTag (TagKind.Image (uri, alt, width, height))   noTags
  let Object tags               = makeTag TagKind.Object                              tags
  let table ths trs             = makeTag (TagKind.Table (ths, trs))                  noTags
  let b tags                    = makeTag (TagKind.CharacterStyle Bold          )     tags
  let em tags                   = makeTag (TagKind.CharacterStyle Emphasized    )     tags
  let strong tags               = makeTag (TagKind.CharacterStyle Strong        )     tags
  let i tags                    = makeTag (TagKind.CharacterStyle Italic        )     tags
  let u tags                    = makeTag (TagKind.CharacterStyle Underline     )     tags
  let strike tags               = makeTag (TagKind.CharacterStyle StrikeThrough )     tags
  let del tags                  = makeTag (TagKind.CharacterStyle Deleted       )     tags
  let sup tags                  = makeTag (TagKind.CharacterStyle Superscript   )     tags
  let sub tags                  = makeTag (TagKind.CharacterStyle Subscript     )     tags
  let cite tags                 = makeTag (TagKind.CharacterStyle Citation      )     tags
  let font f c tags             = makeTag (TagKind.CharacterStyle (Font (f,c))  )     tags
  let text s                    = makeTag (TagKind.Text s)                            noTags

  let setId id (Tag (tagKind, _, styles, tags)) = Tag (tagKind, TagId id, styles, tags)

  let setBackgroundColor c    = liftStyle <| fun ss -> { ss with BackgroundColor = Some c  }
  let setColor           c    = liftStyle <| fun ss -> { ss with Color           = Some c  }
  let setFontFamily      ff   = liftStyle <| fun ss -> { ss with FontFamily      = Some ff }
  let setFontSize        fs   = liftStyle <| fun ss -> { ss with FontSize        = Some fs }
  let setFontStyle       fs   = liftStyle <| fun ss -> { ss with FontStyle       = Some fs }
  let setFontWeight      fw   = liftStyle <| fun ss -> { ss with FontWeight      = Some fw }
  let setMaxWidth        mw   = liftStyle <| fun ss -> { ss with MaxWidth        = Some mw }
  let setMaxHeight       mh   = liftStyle <| fun ss -> { ss with MaxHeight       = Some mh }
  let setTextAlign       ta   = liftStyle <| fun ss -> { ss with TextAlign       = Some ta }
  let setTextDecoration  td   = liftStyle <| fun ss -> { ss with TextDecoration  = Some td }
  let setCustomStyle     cs   = liftStyle <| fun ss -> { ss with CustomStyle     = Some cs }

open ApiEndPoint

let doIt =
  async {
    (*
    let! accessToken, authenticationToken = Authenticate.logon "0000000048156AC9"
    let printable = sprintf "AccessToken: %s\nAuthenticationToken: %s" accessToken authenticationToken
    printfn "%s" printable

    Clipboard.SetText accessToken
    *)

    let accessToken = "EwCIAq1DBAAUGCCXc8wU/zFu9QnLdZXy%2bYnElFkAARPKs/AXFjtUdk2y5mR/aM5nc6suUgcOXhJvQ6k41EWCQ%2bLgnDQrtqHWyzGTvn0SlrVCurGzga9qDRsDp2qL58sxlc1lZKoFvu/fA9rWeoD1f4QMv73DnaStz4HOvg5F535vH7BS6ujG10alByEtDIenSrW7Od4/MQGSlYe4N%2b%2bQfUku%2bQ4ImOzcS64c%2bPM3eoXP/8Ku1jao/j4Fl27T7iYbYpts9s0lt3ED146pjg8/VKv76NvqszMh1SafYx3404IR90QSCaQ9p2MwPGh%2b0/qpNJbHB4%2bCp4XdIbxm%2bwq2tSgrs9oPYGoZapwTAyvUDvZbzUJmnWBjpSWcUcJmkSQDZgAACHFR5K/Wnas1WAHz0APHWG8t5OklA82v0gsMXDGK43cdYJaPQuCeRaACN8QrWVXwptKb8F5IWh9671EBIgIFMdVXAnTR%2bqeppSbIRiyg%2b7aMIBCSYa75I7XC0yKsCO8Blo57SouMcNgW3rULB598Uyh2eG4ubhgdfEPU/eMcLu9OwqkxSlLeaM/OgV6KBM5vLv09CA%2bUfc61ckyY6/adRa69ModxquhWcApX8fzcEC0VlrPz6jRaf0MOnaEye1c3uMoW8e79skKo76m33VqQs3GIL8a8IWL3AzvVJyH3TZ2uQcQH2fetA8Txmt8G47JhUWzcxUWflEn4DYvIp7nZldJPNv%2b1Kh4sxYXtPP0Ev7k8Mq/ZohM5fTnmHTR6V%2bS1OpvynutHQp2ZfvDgyX7u2dp%2bnPYyQrzmkecC1RZwyFP4KO7ewC57g8Y662uFSrFfkc3WNCdc9Lfdz9pRWuTXfmdUDWsB"

    let! notebooks = getNotebooks accessToken
    printfn "Notebooks:\n%s" notebooks

    let! notebook = getNotebook accessToken (NotebookId "0-58ED096273EF8D4!1373")
    printfn "Notebook:\n%s" notebook

    let! sectionGroups = getNotebook_SectionGroups accessToken (NotebookId "0-58ED096273EF8D4!1373")
    printfn "SectionGroups:\n%s" sectionGroups

    let! sections = getNotebook_Sections accessToken (NotebookId "0-58ED096273EF8D4!1373")
    printfn "Sections:\n%s" sections

    (*
    let! pages = getAllPages accessToken
    printfn "Pages:\n%s" pages

    let! page = getPage accessToken (PageId "0-edd4956daf4301413a9270c6ba850ec0!1-58ED096273EF8D4!648")
    printfn "Page:\n%s" page

    let! pageContent = getPageContent accessToken (PageId "0-edd4956daf4301413a9270c6ba850ec0!1-58ED096273EF8D4!648") true
    printfn "PageContent:\n%s" pageContent
    *)
  }

open PageModel

[<STAThread>]
[<EntryPoint>]
let main argv = 

  let p = 
    page "Hello" NoStyle
      [|
        div 
          [|
            img (Uri "http://google.com") "Hello" 100 200
            span [| text "Yello" |]
          |]
      |]

  let ps = renderPage p
  printfn "%s" <| ps

  let previous  = SynchronizationContext.Current
  use current   = new WindowsFormsSynchronizationContext ()
  SynchronizationContext.SetSynchronizationContext current

  try
    try
      let cont = ref true
      let action =
        async {
          do! doIt
          cont := false
        }

      Async.StartImmediate action
    
      while !cont do
        Application.DoEvents ()
        Thread.Sleep 10
      
      0
    with
    | e -> printfn "Caught exception: %s" e.Message; 999
  finally
    SynchronizationContext.SetSynchronizationContext previous
