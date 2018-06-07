<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:fo="http://www.w3.org/1999/XSL/Format" version="1.0">
	<xsl:output encoding="UTF-8" indent="yes" method="xml" standalone="no" omit-xml-declaration="no"/>
	<xsl:template match="body">
		<fo:root>
			<fo:layout-master-set>
				<fo:simple-page-master master-name="A4-portrail" page-height="297mm" page-width="210mm" margin-top="20mm" margin-bottom="25mm" margin-left="20mm" margin-right="20mm">
					<fo:region-body/> <!--margin-top="25mm" margin-bottom="20mm"-->
					<fo:region-before region-name="xsl-region-before" extent="25mm" display-align="before" precedence="true"/>
				</fo:simple-page-master>
			</fo:layout-master-set>
			
			<fo:page-sequence master-reference="A4-portrail">
				<fo:flow flow-name="xsl-region-body" border-collapse="collapse" reference-orientation="0">
					
					
					<fo:table table-layout="fixed"  width="165mm" font-size="10pt" border-color="black" border-width="0.35mm" border-style="solid" text-align="right" display-align="center">
						<fo:table-body>
							<fo:table-row height="10mm">
								<fo:table-cell>
									<fo:block><fo:external-graphic src="E:work/1/image.bmp"  content-height="scale-to-fit" scaling="non-uniform"/></fo:block>
								</fo:table-cell>
							</fo:table-row>
						</fo:table-body>	
					</fo:table>
					
					<fo:table table-layout="fixed" width="165mm" font-size="10pt" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="3mm" >
						<fo:table-column column-width="120mm"/>
						<fo:table-column column-width="20mm"/>
						<fo:table-column column-width="25mm"/>
						<fo:table-body font-size="14px">
							<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>Template Name</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>Version</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>Date</fo:block>
								</fo:table-cell>
							</fo:table-row>
							<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="templateName"/></fo:block>									
								</fo:table-cell>
									<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
								<fo:block><xsl:value-of select="version"/></fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="date"/></fo:block>
								</fo:table-cell>
							</fo:table-row>							
						</fo:table-body>
					</fo:table>

					<fo:table table-layout="fixed" width="165mm" font-size="10pt" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="3mm" >
						<fo:table-column column-width="45mm"/>
						<fo:table-column column-width="40mm"/>
						<fo:table-column column-width="40mm"/>
						<fo:table-column column-width="40mm"/>
						<fo:table-body font-size="14px">
							<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>Main Contractor</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="mainContractor"/></fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>Project Name</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="projectName"/></fo:block>
								</fo:table-cell>
							</fo:table-row>
							<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>Management Company</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="managementCompany"/></fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>Contract No</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="contractNo"/></fo:block>
								</fo:table-cell>
							</fo:table-row>
							<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>QC Company</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="qcCompany"/></fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>QA Company</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="qaCompany"/></fo:block>
								</fo:table-cell>
							</fo:table-row>							
						</fo:table-body>
					</fo:table>					
					
					<fo:table table-layout="fixed" width="165mm" font-size="10pt" border-color="black" border-width="0.5mm" border-style="solid" text-align="center" display-align="center" space-after="3mm" >
						<fo:table-column column-width="85mm"/>
						<fo:table-column column-width="80mm"/>
						<fo:table-body font-size="14px">
							<fo:table-row height="7mm">
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block>NCR Number</fo:block>
								</fo:table-cell>
								<fo:table-cell border-color="black" border-width="0.175mm" border-style="solid">
									<fo:block><xsl:value-of select="ncrNumber"/></fo:block>
								</fo:table-cell>	
							</fo:table-row>
						</fo:table-body>
					</fo:table>

					
					
					<fo:table table-layout="fixed" width="100%" font-size="10pt" border-color="black" border-width="0.35mm" border-style="solid" text-align="center" display-align="center" space-after="5mm">
						<fo:table-column column-width="proportional-column-width(20)"/>
						<fo:table-column column-width="proportional-column-width(30)"/>
						<fo:table-column column-width="proportional-column-width(25)"/>
						<fo:table-column column-width="proportional-column-width(50)"/>
						<fo:table-body font-size="95%">
							<fo:table-row height="8mm">
								<fo:table-cell>
									<fo:block>Cod.Cliente</fo:block>
								</fo:table-cell>
								<fo:table-cell>
									<fo:block>Ragione Sociale</fo:block>
								</fo:table-cell>
								<fo:table-cell>
									<fo:block>Codice Fiscale/P.IVA</fo:block>
								</fo:table-cell>
								<fo:table-cell>
									<fo:block>Indirizzo</fo:block>
								</fo:table-cell>
							</fo:table-row>
							<xsl:for-each select="record">
								<fo:table-row>
									<fo:table-cell>
										<fo:block>
											<xsl:value-of select="codice_cliente"/>
										</fo:block>
									</fo:table-cell>
									<fo:table-cell>
										<fo:block>
											<xsl:value-of select="rag_soc"/>
										</fo:block>
									</fo:table-cell>
									<fo:table-cell>
										<fo:block>
											<xsl:value-of select="codice_fiscale"/>
										</fo:block>
									</fo:table-cell>
									<fo:table-cell>
										<fo:block>
											<xsl:value-of select="indirizzo"/>
										</fo:block>
									</fo:table-cell>
								</fo:table-row>
							</xsl:for-each>
						</fo:table-body>
					</fo:table>

					
					<fo:block id="end-of-document">
					</fo:block>
				</fo:flow>
			</fo:page-sequence>
		</fo:root>
	</xsl:template>
</xsl:stylesheet>
