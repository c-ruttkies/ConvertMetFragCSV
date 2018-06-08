package de.ipbhalle.main;

import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.Rectangle;
import java.awt.image.BufferedImage;
import java.awt.image.RenderedImage;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.ArrayList;
import java.util.Map;

import javax.imageio.ImageIO;

import org.openscience.cdk.AtomContainer;
import org.openscience.cdk.DefaultChemObjectBuilder;
import org.openscience.cdk.aromaticity.Aromaticity;
import org.openscience.cdk.aromaticity.ElectronDonation;
import org.openscience.cdk.exception.CDKException;
import org.openscience.cdk.exception.InvalidSmilesException;
import org.openscience.cdk.graph.Cycles;
import org.openscience.cdk.inchi.InChIGeneratorFactory;
import org.openscience.cdk.inchi.InChIToStructure;
import org.openscience.cdk.interfaces.IAtomContainer;
import org.openscience.cdk.layout.StructureDiagramGenerator;
import org.openscience.cdk.renderer.AtomContainerRenderer;
import org.openscience.cdk.renderer.RendererModel;
import org.openscience.cdk.renderer.font.AWTFontManager;
import org.openscience.cdk.renderer.generators.BasicAtomGenerator;
import org.openscience.cdk.renderer.generators.BasicAtomGenerator.AtomRadius;
import org.openscience.cdk.renderer.generators.BasicBondGenerator;
import org.openscience.cdk.renderer.generators.BasicSceneGenerator;
import org.openscience.cdk.renderer.generators.IGenerator;
import org.openscience.cdk.renderer.visitor.AWTDrawVisitor;
import org.openscience.cdk.smiles.SmilesParser;
import org.openscience.cdk.tools.manipulator.AtomContainerManipulator;
import org.openscience.cdk.silent.SilentChemObjectBuilder;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableImage;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;
import net.sf.jniinchi.INCHI_RET;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

public class ConvertMetFragCSV {

	public static String fileSep = System.getProperty("file.separator");
	public static String csvFile = "";
	public static String resultspath = "";
	public static String fileName = "";
	public static boolean withImages = false;
	public static Map<String, Integer> labels = new HashMap<String, Integer>();
	public static java.util.Vector<String> skipWhenMissing;
	public static java.util.Vector<String> propertyWhiteList;
	public static String sheetName = "Candidate List";
	public static String moleculeContainerProperty = "InChI"; //inchi
	public static String format = "xls";
	public static String urlName = "MetFragWebURL";

	public static InChIGeneratorFactory inchiFactory;
	
	/**
	 * 
	 * @param args
	 */
	public static void main(String[] args) {
		if (args.length <= 2) {
			System.out.println("usage: command csv='csv file' output='output file' "
					+ "propertyWhiteList=PROPERTY1[,PROPERTY2,...] [format='tex or xls']"
					+ " [img='xls with images?'] [sheetName=SHEETNAME] [moleculeContainerProperty=MoleculeContainerProperty]");
			System.exit(1);
		}

		String arg_string = args[0];
		for (int i = 1; i < args.length; i++) {
			arg_string += " " + args[i];
		}
		arg_string = arg_string.replaceAll("\\s*=\\s*", "=").trim();

		String[] args_spaced = arg_string.split("\\s+");
		for (int i = 0; i < args_spaced.length; i++) {
			String[] tmp = args_spaced[i].split("=");
			if (tmp[0].equals("csv"))
				csvFile = tmp[1];
			else if (tmp[0].equals("output"))
				resultspath = tmp[1];
			else if (tmp[0].equals("format"))
				format = tmp[1];
			else if (tmp[0].equals("sheetName"))
				sheetName = tmp[1];
			else if (tmp[0].equals("moleculeContainerProperty"))
				moleculeContainerProperty = tmp[1];
			else if (tmp[0].equals("img") && tmp[1].charAt(0) == '1')
				withImages = true;
		    else if (tmp[0].equals("propertyWhiteList")) {
				String[] propertyNamesWhiteList = tmp[1].split(",");
				propertyWhiteList = new java.util.Vector<String>();
				for (String property : propertyNamesWhiteList)
					propertyWhiteList.add(property);
			} else {
				System.err.println("Parameter unknown " + args_spaced[i]);
				System.exit(1);
			}
		}
		
		if(propertyWhiteList == null) {
			System.out.println("Please provide a white list for properties.");
			System.exit(1);
		}
		
		try {
			inchiFactory = InChIGeneratorFactory.getInstance();
		} catch (CDKException e2) { 
			e2.printStackTrace();
			System.exit(1);
		}
		
		// get file reader for the sdf file
		File file = new File(csvFile);

		try {
			if (!file.getName().endsWith(".csv")) {
				System.out.println("csv file extension missing");
				throw new Exception();
			}
		} catch (FileNotFoundException e) {
			System.err.println("Could not read sdf file. Is it valid?");
			System.exit(1);
		} catch (Exception e) {
			System.err.println("Valid sdf file?");
			System.exit(1);
		}
		
		try {
			if(format.equals("xls")) writeXLSFile(file);
			else if(format.equals("tex")) writeTexFile(file);
			else System.err.println("Format " + format + " not known. Specify 'tex' or 'xls'!");
		} catch (CloneNotSupportedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
			System.exit(1);
		}

		// write xls file
		
	}

	/**
	 * writes out xls file containing all molecules of containersList
	 * 
	 * @param containersList
	 * @throws CloneNotSupportedException
	 * @throws IOException 
	 */
	public static void writeXLSFile(File csvFile) throws CloneNotSupportedException, IOException {

		File xlsFile = new File(resultspath);
		WritableSheet sheet = null;
		WritableWorkbook workbook = null;

		try {
			xlsFile.createNewFile();
			workbook = Workbook.createWorkbook(xlsFile);
		} catch (IOException e) {
			System.out.println("Could not create xls file.");
			System.exit(1);
		}

		sheet = workbook.createSheet(sheetName, 0);

		java.io.Reader reader = new java.io.InputStreamReader(new java.io.FileInputStream(csvFile), Charset.forName("UTF-8"));
		CSVParser csvFileParser = new CSVParser(reader, CSVFormat.EXCEL.withHeader());
		
		java.util.Iterator<CSVRecord> it = csvFileParser.iterator();

		List<IAtomContainer> containersList = new ArrayList<IAtomContainer>();
		
		if(!csvFileParser.getHeaderMap().containsKey(moleculeContainerProperty)) {
			System.err.println("CSV does not contain moleculeContainerProperty " + moleculeContainerProperty + " as a field");
			System.exit(1);
		}
		
		while(it.hasNext()) {
			CSVRecord record = it.next();
			IAtomContainer con = retrieveAtomContainer(record.get(moleculeContainerProperty));
			for(int i = 0;  i < propertyWhiteList.size(); i++) {
				String propValue = record.get(propertyWhiteList.get(i));
				if(propValue == null) propValue = "";
				con.setProperty(propertyWhiteList.get(i), propValue);
			}
			containersList.add(con);
		}
		csvFileParser.close();
		// add images if selected
		int columnWidthAdd = withImages ? 3 : 0;
		int rowHeightAdd = withImages ? 9 : 1;
		List<RenderedImage> molImages = null;
		if (withImages) {
			
			molImages = convertMoleculesToImages(containersList);
			for (int i = 0; i < molImages.size(); i++) {
				// File imageFile = new File(resultspath + fileSep + fileName+
				// "_" +i+".png");
				try {
					File imageFile = File.createTempFile("image", ".png", new File(System.getProperty("java.io.tmpdir")));
					imageFile.deleteOnExit();
					if (ImageIO.write(molImages.get(i), "png", imageFile)) {
						WritableImage wi = new WritableImage(0, (i * rowHeightAdd) + 1, columnWidthAdd, rowHeightAdd,
								imageFile);
						sheet.addImage(wi);
					}
				} catch (IOException e) {
					e.printStackTrace();
				}

			}
		}

		WritableFont arial10font = new WritableFont(WritableFont.ARIAL, 10);
		WritableCellFormat arial10format = new WritableCellFormat(arial10font);
		try {
			arial10font.setBoldStyle(WritableFont.BOLD);
		} catch (WriteException e1) {
			System.out.println("Warning: Could not set WritableFont");
		}

		int numberCells = 0;
		for (int i = 0; i < containersList.size(); i++) {
			// write header
			Map<Object, Object> molProperties = containersList.get(i).getProperties();
			Iterator<Object> propNames = molProperties.keySet().iterator();
			// just in case images are used
			int column = 3;
			int row = (i * rowHeightAdd) + 1;
			int currentPropertyIndex = 0;
			while (propNames.hasNext()) {
				String propName = (String) propNames.next();
				if(!withImages) {
					if (!labels.containsKey(propName)) {
						labels.put(propName, new Integer(numberCells));
						try {
							sheet.addCell(new Label(labels.get(propName) + columnWidthAdd, 0, propName, arial10format));
						} catch (RowsExceededException e) {
							e.printStackTrace();
						} catch (WriteException e) {
							e.printStackTrace();
						}
						numberCells++;
					}
					try {
						sheet.addCell(new Label(labels.get(propName) + columnWidthAdd, (i * rowHeightAdd) + 1,
								(String) molProperties.get(propName)));
					} catch (RowsExceededException e) {
						e.printStackTrace();
					} catch (WriteException e) {
						e.printStackTrace();
					}
				} else {
					try {
						sheet.addCell(new Label(column, row + currentPropertyIndex, propName, arial10format));
						sheet.addCell(new Label(column + 1, row + currentPropertyIndex, (String) molProperties.get(propName)));
					} catch (RowsExceededException e) {
						e.printStackTrace();
						currentPropertyIndex++;
						continue;
					} catch (WriteException e) {
						e.printStackTrace();
						currentPropertyIndex++;
						continue;
					}
					currentPropertyIndex++;
					// only write n rows
					if(currentPropertyIndex == 6) {
						currentPropertyIndex = 0;
						column += 2;
					}
				}
			}
		}

		try {
			workbook.write();
			workbook.close();
			System.out.println("Wrote xls file to " + xlsFile.getAbsolutePath());
		} catch (IOException e) {
			System.out.println("Could not write xls file.");
			e.printStackTrace();
			System.exit(1);
		}

		System.out.println("Read " + containersList.size() + " molecules");
	}

	public static void writeTexFile(File csvFile) throws CloneNotSupportedException, IOException {
		java.util.Vector<String> texLines = new java.util.Vector<String>();
		
		java.io.Reader reader = new java.io.InputStreamReader(new java.io.FileInputStream(csvFile), Charset.forName("UTF-8"));
		CSVParser csvFileParser = new CSVParser(reader, CSVFormat.EXCEL.withHeader());
		
		java.util.Iterator<CSVRecord> it = csvFileParser.iterator();

		List<IAtomContainer> containersList = new ArrayList<IAtomContainer>();
		
		if(!csvFileParser.getHeaderMap().containsKey(moleculeContainerProperty)) {
			System.err.println("CSV does not contain moleculeContainerProperty " + moleculeContainerProperty + " as a field");
			System.exit(1);
		}
		
		addTexHeaders(texLines);
		
		while(it.hasNext()) {
			CSVRecord record = it.next();
			IAtomContainer con = retrieveAtomContainer(record.get(moleculeContainerProperty));
			for(int i = 0;  i < propertyWhiteList.size(); i++) {
				String propValue = "";
				if(record.isMapped(propertyWhiteList.get(i))) {
					propValue = record.get(propertyWhiteList.get(i));
					if(propValue == null) propValue = "";
				} 
				con.setProperty(propertyWhiteList.get(i), propValue);
			}
			containersList.add(con);
		}
		csvFileParser.close();
		// add images if selected
		List<RenderedImage> molImages = convertMoleculesToImages(containersList);
	
		for (int i = 0; i < containersList.size(); i++) {
			File imageFile = File.createTempFile("image", ".png", new File(System.getProperty("java.io.tmpdir")));
			
			texLines.add("\\begin{minipage}{1\\textwidth}");
			texLines.add("	\\begin{minipage}{0.15\\textwidth}");
			
			if (ImageIO.write(molImages.get(i), "png", imageFile)) {
				texLines.add("		\\includegraphics[scale=0.5]{" + imageFile + "}");
			}
			texLines.add("	\\end{minipage} \\hfill");
			texLines.add("	\\begin{minipage}{0.8\\textwidth}");
			texLines.add("		\\begin{itemize}");
			Map<Object, Object> molProperties = containersList.get(i).getProperties();
			Iterator<Object> propNames = molProperties.keySet().iterator();
			// just in case images are used
			String urlname = "";
			while(propNames.hasNext()) {
				String propName = (String)propNames.next();
				String propValue = (String)molProperties.get(propName);
				propValue = propValue.replaceAll("\\$", "").replaceAll("\\^", "").replaceAll("%", "").replaceAll("_", "\\\\_").trim();
				if(!propName.equals(urlName) && propValue.length() != 0) {
					texLines.add("			\\item[] \\textbf{" + propName + "} " + propValue);
				} else {
					urlname = propValue;
				}
			}
			texLines.add("		\\end{itemize}");
			texLines.add("	\\end{minipage}\\\\[0.4cm]");
			if(!urlname.equals("")) {
				texLines.add("\\textbf{MetFragWeb:} \\href{" + urlname + "}{Send query to MetFragWeb}");
			}
			texLines.add("\\end{minipage}\\\\[0.8cm]");
			texLines.add("");
		}
		
		java.io.BufferedWriter bwriter = new java.io.BufferedWriter(new java.io.FileWriter(new java.io.File(resultspath)));
	
		addTexFooters(texLines);
		
		for(int i = 0; i < texLines.size(); i++) {
			bwriter.write(texLines.get(i));
			bwriter.newLine();
		}
		bwriter.close();
		System.out.println("Read " + containersList.size() + " molecules");
	}
	
	public static void addTexHeaders(java.util.Vector<String> texlines) {
		texlines.add("\\documentclass[9pt]{article}");
		texlines.add("\\usepackage[T1]{fontenc}");
		texlines.add("\\usepackage[english]{babel}");
		texlines.add("\\usepackage{lmodern}");
		texlines.add("\\usepackage{tabularx}");
		texlines.add("\\usepackage[margin=0.5in]{geometry}");
		//texlines.add("\\usepackage{float}");
		texlines.add("\\usepackage{hyperref}");
		texlines.add("\\usepackage{graphicx}");
		texlines.add("\\begin{document}");
		texlines.add("\\setlength\\parindent{0pt}");
	}

	public static void addTexFooters(java.util.Vector<String> texlines) {
		texlines.add("\\end{document}");
	}
	
	public static IAtomContainer retrieveAtomContainer(String moleculeString) {
		if(moleculeContainerProperty.toLowerCase().equals("smiles")) return parseSmiles(moleculeString);
		else if(moleculeContainerProperty.toLowerCase().equals("inchi"))
			try {
				return getAtomContainerFromInChI(moleculeString);
			} catch (Exception e) {
				return new AtomContainer();
			}
		else return new AtomContainer();
	}

	public static IAtomContainer getAtomContainerFromInChI(String inchi) throws Exception {
		InChIToStructure its = inchiFactory.getInChIToStructure(inchi, DefaultChemObjectBuilder.getInstance());
		if(its == null) {
			throw new Exception("InChI problem: " + inchi);
		}
		INCHI_RET ret = its.getReturnStatus();
		if (ret == INCHI_RET.WARNING) {
		//	logger.warn("InChI warning: " + its.getMessage());
		} else if (ret != INCHI_RET.OKAY) {
			throw new Exception("InChI problem: " + inchi);
		}
		IAtomContainer molecule = its.getAtomContainer();
		try {
			AtomContainerManipulator.percieveAtomTypesAndConfigureAtoms(molecule);
			Aromaticity arom = new Aromaticity(ElectronDonation.cdk(), Cycles.cdkAromaticSet());
			arom.apply(molecule);
		} catch (CDKException e) {
			e.printStackTrace();
		}
		return molecule;
	}
	
	public static IAtomContainer parseSmiles(String smiles) {
		SmilesParser sp = new SmilesParser(SilentChemObjectBuilder.getInstance());
		IAtomContainer precursorMolecule = null;
		try {
			precursorMolecule = sp.parseSmiles(smiles);
		} catch (InvalidSmilesException e) {
			e.printStackTrace();
		}
		return precursorMolecule;
	}
	
	/**
	 * generate images of chemical structures
	 * 
	 * @param mol
	 * @return
	 * @throws CloneNotSupportedException
	 * @throws Exception
	 */
	private static List<RenderedImage> convertMoleculesToImages(List<IAtomContainer> mols)
			throws CloneNotSupportedException {

		List<RenderedImage> molImages = new ArrayList<RenderedImage>();

		int width = 200;
		int height = 200;

		for (int i = 0; i < mols.size(); i++) {
			IAtomContainer mol = AtomContainerManipulator.removeHydrogens(mols.get(i));
			IAtomContainer molSource = mol.clone();

			try {
				AtomContainerManipulator.percieveAtomTypesAndConfigureAtoms(molSource);
			} catch (CDKException e1) {
				e1.printStackTrace();
			}
			Rectangle drawArea = new Rectangle(width, height);
			Image image = new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB);

			StructureDiagramGenerator sdg = new StructureDiagramGenerator();
			sdg.setMolecule(molSource);
			try {
				sdg.generateCoordinates();
			} catch (Exception e) {
				System.out.println("Warning: Could not draw molecule number " + (i + 1) + ".");
				molImages.add(new BufferedImage(width, height, BufferedImage.TYPE_INT_RGB));
				continue;
			}
			molSource = sdg.getMolecule();

			List<IGenerator<IAtomContainer>> generators = new ArrayList<IGenerator<IAtomContainer>>();
			generators.add(new BasicSceneGenerator());
			generators.add(new BasicBondGenerator());
			generators.add(new BasicAtomGenerator());

			AtomContainerRenderer renderer = new AtomContainerRenderer(generators, new AWTFontManager());
			RendererModel rm = renderer.getRenderer2DModel();
			rm.set(AtomRadius.class, 0.4);

			renderer.setup(molSource, drawArea);

			Graphics2D g2 = (Graphics2D) image.getGraphics();
			g2.setColor(Color.WHITE);
			g2.fillRect(0, 0, width, height);

			renderer.paint(molSource, new AWTDrawVisitor(g2), drawArea, true);

			molImages.add((RenderedImage) image);

		}
		return molImages;
	}

}
