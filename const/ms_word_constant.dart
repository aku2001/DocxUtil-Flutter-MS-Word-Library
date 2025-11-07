// width, height, imageId
const String msImageTemplate = """
  <w:p w14:paraId="2FE9FB91" w14:textId="693882D8"
      w:rsidR="00DD6D66" w:rsidRDefault="00824F69"
      w:rsidP="00DD13FA">
      <w:pPr>
          <w:jc w:val="center" />
      </w:pPr>
      <w:r>
          <w:rPr>
              <w:noProof />
          </w:rPr>
          <w:drawing>
              <wp:inline distT="0" distB="0"
                  distL="0" distR="0"
                  wp14:anchorId="38007831"
                  wp14:editId="050FB82F">
                  <wp:extent cx="{width}"
                      cy="{height}" />
                  <wp:effectExtent l="0" t="0"
                      r="0" b="0" />
                  <wp:docPr id="344881166"
                      name="Resim 2"
                      descr="metin, ekran görüntüsü, yazılım içeren bir resim&#xA;&#xA;Yapay zeka tarafından oluşturulan içerik yanlış olabilir." />
                  <wp:cNvGraphicFramePr>
                      <a:graphicFrameLocks
                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                          noChangeAspect="1" />
                  </wp:cNvGraphicFramePr>
                  <a:graphic
                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                      <a:graphicData
                          uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                          <pic:pic
                              xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                              <pic:nvPicPr>
                                  <pic:cNvPr
                                      id="344881166"
                                      name="Resim 2"
                                      descr="metin, ekran görüntüsü, yazılım içeren bir resim&#xA;&#xA;Yapay zeka tarafından oluşturulan içerik yanlış olabilir." />
                                  <pic:cNvPicPr>
                                      <a:picLocks
                                          noChangeAspect="1"
                                          noChangeArrowheads="1" />
                                  </pic:cNvPicPr>
                              </pic:nvPicPr>
                              <pic:blipFill>
                                  <a:blip
                                      r:embed="{imageId}"
                                      cstate="print">
                                      <a:extLst>
                                          <a:ext
                                              uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                              <a14:useLocalDpi
                                                  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                                                  val="0" />
                                          </a:ext>
                                      </a:extLst>
                                  </a:blip>
                                  <a:srcRect />
                                  <a:stretch>
                                      <a:fillRect />
                                  </a:stretch>
                              </pic:blipFill>
                              <pic:spPr
                                  bwMode="auto">
                                  <a:xfrm>
                                      <a:off x="0"
                                          y="0" />
                                      <a:ext
                                          cx="{width}"
                                          cy="{height}" />
                                  </a:xfrm>
                                  <a:prstGeom
                                      prst="rect">
                                      <a:avLst />
                                  </a:prstGeom>
                                  <a:noFill />
                                  <a:ln>
                                      <a:noFill />
                                  </a:ln>
                              </pic:spPr>
                          </pic:pic>
                      </a:graphicData>
                  </a:graphic>
              </wp:inline>
          </w:drawing>
      </w:r>
  </w:p>

""";

const String msPageBreak = """

  <w:p>
      <w:r>
          <w:br w:type="page" />
      </w:r>
  </w:p>

""";

const String msLineBreak = """

  <w:p>
      <w:r>
          <w:t></w:t>
      </w:r>
  </w:p>

""";

// imageIdNumber, imageName
const String msRelationShipTemplate = """
  <Relationship Id="rId{imageIdNumber}"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    Target="media/{imageName}" />  

""";

// textBoxWidth, textBoxHeight, imageWidth, imageHeight, imageId
const String msTextBoxImageTemplate = """

  <w:p w14:paraId="1384BDCE" w14:textId="31BD7554" w:rsidR="00FB4338" w:rsidRPr="00BA78E1"
      w:rsidRDefault="00EC5553" w:rsidP="00FB4338">
      <w:pPr>
          <w:jc w:val="center" />
          <w:rPr>
              <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
              <w:sz w:val="22" />
              <w:szCs w:val="22" />
          </w:rPr>
      </w:pPr>
      <w:r w:rsidRPr="00EC5553">
          <w:rPr>
              <w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:cs="Arial" />
              <w:noProof />
              <w:sz w:val="22" />
              <w:szCs w:val="22" />
          </w:rPr>
          <mc:AlternateContent>
              <mc:Choice Requires="wps">
                  <w:drawing>
                      <wp:anchor distT="0" distB="0" distL="0" distR="0"
                          simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0"
                          layoutInCell="1" allowOverlap="1" wp14:anchorId="4C4AF7A6"
                          wp14:editId="7045A84D">
                          <wp:simplePos x="0" y="0" />
                          <wp:positionH relativeFrom="margin">
                              <wp:align>right</wp:align>
                          </wp:positionH>
                          <wp:positionV relativeFrom="paragraph">
                              <wp:posOffset>0</wp:posOffset>
                          </wp:positionV>
                          <wp:extent cx="{textBoxWidth}" cy="{textBoxHeight}" />
                          <wp:effectExtent l="0" t="0" r="0" b="0" />
                          <wp:wrapNone />
                          <wp:docPr id="217" name="Metin Kutusu 2" />
                          <wp:cNvGraphicFramePr>
                              <a:graphicFrameLocks
                                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" />
                          </wp:cNvGraphicFramePr>
                          <a:graphic
                              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                              <a:graphicData
                                  uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                                  <wps:wsp>
                                      <wps:cNvSpPr txBox="1">
                                          <a:spLocks noChangeArrowheads="1" />
                                      </wps:cNvSpPr>
                                      <wps:spPr bwMode="auto">
                                          <a:xfrm>
                                              <a:off x="0" y="0" />
                                              <a:ext cx="{textBoxWidth}" cy="{textBoxHeight}" />
                                          </a:xfrm>
                                          <a:prstGeom prst="rect">
                                              <a:avLst />
                                          </a:prstGeom>
                                          <a:noFill />
                                          <a:ln w="9525">
                                              <a:noFill />
                                              <a:miter lim="800000" />
                                              <a:headEnd />
                                              <a:tailEnd />
                                          </a:ln>
                                      </wps:spPr>
                                      <wps:txbx>
                                          <w:txbxContent>
                                              <w:p w14:paraId="5FBB5BCB" w14:textId="6558C19C"
                                                  w:rsidR="00EC5553" w:rsidRDefault="00EC5553"
                                                  w:rsidP="00EC5553">
                                                  <w:pPr>
                                                      <w:jc w:val="center" />
                                                  </w:pPr>
                                                  <w:r>
                                                      <w:rPr>
                                                          <w:noProof />
                                                      </w:rPr>
                                                      <w:drawing>
                                                          <wp:inline distT="0" distB="0"
                                                              distL="0" distR="0"
                                                              wp14:anchorId="4DB787E0"
                                                              wp14:editId="269D96C2">
                                                              <wp:extent cx="{imageWidth}"
                                                                  cy="{imageHeight}" />
                                                              <wp:effectExtent l="0" t="0"
                                                                  r="0" b="0" />
                                                              <wp:docPr id="241413860"
                                                                  name="Resim 2"
                                                                  descr="grafik, yazı tipi, logo, grafik tasarım içeren bir resim&#xA;&#xA;Yapay zeka tarafından oluşturulan içerik yanlış olabilir." />
                                                              <wp:cNvGraphicFramePr>
                                                                  <a:graphicFrameLocks
                                                                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                                                      noChangeAspect="1" />
                                                              </wp:cNvGraphicFramePr>
                                                              <a:graphic
                                                                  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                                                  <a:graphicData
                                                                      uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                                      <pic:pic
                                                                          xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                                          <pic:nvPicPr>
                                                                              <pic:cNvPr
                                                                                  id="241413860"
                                                                                  name="Resim 2"
                                                                                  descr="grafik, yazı tipi, logo, grafik tasarım içeren bir resim&#xA;&#xA;Yapay zeka tarafından oluşturulan içerik yanlış olabilir." />
                                                                              <pic:cNvPicPr>
                                                                                  <a:picLocks
                                                                                      noChangeAspect="1"
                                                                                      noChangeArrowheads="1" />
                                                                              </pic:cNvPicPr>
                                                                          </pic:nvPicPr>
                                                                          <pic:blipFill>
                                                                              <a:blip
                                                                                  r:embed="{imageId}">
                                                                                  <a:extLst>
                                                                                      <a:ext
                                                                                          uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                                                          <a14:useLocalDpi
                                                                                              xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                                                                                              val="0" />
                                                                                      </a:ext>
                                                                                  </a:extLst>
                                                                              </a:blip>
                                                                              <a:srcRect />
                                                                              <a:stretch>
                                                                                  <a:fillRect />
                                                                              </a:stretch>
                                                                          </pic:blipFill>
                                                                          <pic:spPr
                                                                              bwMode="auto">
                                                                              <a:xfrm>
                                                                                  <a:off x="0"
                                                                                      y="0" />
                                                                                  <a:ext
                                                                                      cx="{imageWidth}"
                                                                                      cy="{imageHeight}" />
                                                                              </a:xfrm>
                                                                              <a:prstGeom
                                                                                  prst="rect">
                                                                                  <a:avLst />
                                                                              </a:prstGeom>
                                                                              <a:noFill />
                                                                              <a:ln>
                                                                                  <a:noFill />
                                                                              </a:ln>
                                                                          </pic:spPr>
                                                                      </pic:pic>
                                                                  </a:graphicData>
                                                              </a:graphic>
                                                          </wp:inline>
                                                      </w:drawing>
                                                  </w:r>
                                              </w:p>
                                          </w:txbxContent>
                                      </wps:txbx>
                                      <wps:bodyPr rot="0" vert="horz" wrap="square"
                                          lIns="0" tIns="0" rIns="0" bIns="0"
                                          anchor="t" anchorCtr="0">
                                          <a:noAutofit />
                                      </wps:bodyPr>
                                  </wps:wsp>
                              </a:graphicData>
                          </a:graphic>
                          <wp14:sizeRelH relativeFrom="margin">
                              <wp14:pctWidth>0</wp14:pctWidth>
                          </wp14:sizeRelH>
                          <wp14:sizeRelV relativeFrom="margin">
                              <wp14:pctHeight>0</wp14:pctHeight>
                          </wp14:sizeRelV>
                      </wp:anchor>
                  </w:drawing>
              </mc:Choice>
              <mc:Fallback>
                  <w:pict>
                      <v:shapetype w14:anchorId="4C4AF7A6" id="_x0000_t202"
                          coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">
                          <v:stroke joinstyle="miter" />
                          <v:path gradientshapeok="t" o:connecttype="rect" />
                      </v:shapetype>
                      <v:shape id="Metin Kutusu 2" o:spid="_x0000_s1026" type="#_x0000_t202"
                          style="position:absolute;left:0;text-align:left;margin-left:0pt;margin-top:0pt;width:0pt;height:0pt;z-index:251659264;visibility:visible;mso-wrap-style:square;mso-width-percent:0;mso-height-percent:0;mso-wrap-distance-left:9pt;mso-wrap-distance-top:3.6pt;mso-wrap-distance-right:9pt;mso-wrap-distance-bottom:3.6pt;mso-position-horizontal:right;mso-position-horizontal-relative:margin;mso-position-vertical:absolute;mso-position-vertical-relative:text;mso-width-percent:0;mso-height-percent:0;mso-width-relative:margin;mso-height-relative:margin;v-text-anchor:top"
                          o:gfxdata="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&#xA;"
                          filled="f" stroked="f">
                          <v:textbox>
                              <w:txbxContent>
                                  <w:p w14:paraId="5FBB5BCB" w14:textId="6558C19C"
                                      w:rsidR="00EC5553" w:rsidRDefault="00EC5553"
                                      w:rsidP="00EC5553">
                                      <w:pPr>
                                          <w:jc w:val="center" />
                                      </w:pPr>
                                      <w:r>
                                          <w:rPr>
                                              <w:noProof />
                                          </w:rPr>
                                          <w:drawing>
                                              <wp:inline distT="0" distB="0" distL="0"
                                                  distR="0" wp14:anchorId="4DB787E0"
                                                  wp14:editId="269D96C2">
                                                  <wp:extent cx="{imageWidth}" cy="{imageHeight}" />
                                                  <wp:effectExtent l="0" t="0" r="0" b="0" />
                                                  <wp:docPr id="241413860" name="Resim 2"
                                                      descr="grafik, yazı tipi, logo, grafik tasarım içeren bir resim&#xA;&#xA;Yapay zeka tarafından oluşturulan içerik yanlış olabilir." />
                                                  <wp:cNvGraphicFramePr>
                                                      <a:graphicFrameLocks
                                                          xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
                                                          noChangeAspect="1" />
                                                  </wp:cNvGraphicFramePr>
                                                  <a:graphic
                                                      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
                                                      <a:graphicData
                                                          uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                          <pic:pic
                                                              xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                                                              <pic:nvPicPr>
                                                                  <pic:cNvPr id="241413860"
                                                                      name="Resim 2"
                                                                      descr="grafik, yazı tipi, logo, grafik tasarım içeren bir resim&#xA;&#xA;Yapay zeka tarafından oluşturulan içerik yanlış olabilir." />
                                                                  <pic:cNvPicPr>
                                                                      <a:picLocks
                                                                          noChangeAspect="1"
                                                                          noChangeArrowheads="1" />
                                                                  </pic:cNvPicPr>
                                                              </pic:nvPicPr>
                                                              <pic:blipFill>
                                                                  <a:blip r:embed="{imageId}">
                                                                      <a:extLst>
                                                                          <a:ext
                                                                              uri="{28A0092B-C50C-407E-A947-70E740481C1C}">
                                                                              <a14:useLocalDpi
                                                                                  xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"
                                                                                  val="0" />
                                                                          </a:ext>
                                                                      </a:extLst>
                                                                  </a:blip>
                                                                  <a:srcRect />
                                                                  <a:stretch>
                                                                      <a:fillRect />
                                                                  </a:stretch>
                                                              </pic:blipFill>
                                                              <pic:spPr bwMode="auto">
                                                                  <a:xfrm>
                                                                      <a:off x="0" y="0" />
                                                                      <a:ext cx="{imageWidth}"
                                                                          cy="{imageHeight}" />
                                                                  </a:xfrm>
                                                                  <a:prstGeom prst="rect">
                                                                      <a:avLst />
                                                                  </a:prstGeom>
                                                                  <a:noFill />
                                                                  <a:ln>
                                                                      <a:noFill />
                                                                  </a:ln>
                                                              </pic:spPr>
                                                          </pic:pic>
                                                      </a:graphicData>
                                                  </a:graphic>
                                              </wp:inline>
                                          </w:drawing>
                                      </w:r>
                                  </w:p>
                              </w:txbxContent>
                          </v:textbox>
                          <w10:wrap anchorx="margin" />
                      </v:shape>
                  </w:pict>
              </mc:Fallback>
          </mc:AlternateContent>
      </w:r>
  </w:p>

""";

// insertText
const String msTextTemplate = """
<w:p w14:paraId="270C5F77" w14:textId="77777777" w:rsidR="00B130EA"
    w:rsidRDefault="00B130EA" w:rsidP="00B130EA">
  <w:pPr>
    <w:jc w:val="{alignment}" />
    <w:tabs>
      <w:tab w:val="left" w:pos="360" />
      <w:tab w:val="left" w:pos="420" />
    </w:tabs>
    <w:rPr>
      <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" />
      {boldTag}
      {italicTag}
      {underlineTag}
      <w:sz w:val="{fontSize}" />
      <w:szCs w:val="{fontSize}" />
    </w:rPr>
  </w:pPr>
  <w:r>
    <w:rPr>
      <w:rFonts w:ascii="{font}" w:hAnsi="{font}" w:cs="{font}" />
      {boldTag}
      {italicTag}
      {underlineTag}
      <w:sz w:val="{fontSize}" />
      <w:szCs w:val="{fontSize}" />
    </w:rPr>
    <w:t xml:space="preserve">{insertText}</w:t>
  </w:r>
</w:p>
""";