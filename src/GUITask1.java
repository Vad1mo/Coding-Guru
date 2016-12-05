import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JButton;
import javax.swing.JTabbedPane;
import javax.swing.JLabel;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import java.awt.Font;
import javax.swing.JTextArea;
import javax.swing.border.TitledBorder;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.swing.JCheckBox;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.Random;
import java.awt.event.ActionEvent;

public class GUITask1 extends JFrame {

	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	private ArrayList<String> sentences;
	private ArrayList<String> temp;
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				ListCreator creator=new ListCreator();
				creator.LoadAndCreatLists("");
				try {
					try {
						UIManager.setLookAndFeel("com.jtattoo.plaf.bernstein.BernsteinLookAndFeel");
					} catch (ClassNotFoundException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (InstantiationException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (IllegalAccessException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					} catch (UnsupportedLookAndFeelException e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
					GUITask1 frame = new GUITask1();
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public GUITask1() {
	
		setTitle("Task1");
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 583, 495);
		setMaximumSize(this.getPreferredSize());
		setResizable(false);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		sentences=readSentences();
		temp=new ArrayList<>();
		temp.addAll(sentences);
		JLabel lblTask = new JLabel("Select appropriate aspect relevent to the following comments:");
		lblTask.setFont(new Font("Dialog", Font.BOLD, 14));
		lblTask.setHorizontalAlignment(SwingConstants.CENTER);
		lblTask.setBounds(12, 0, 549, 38);
		contentPane.add(lblTask);
		
		JTextArea textArea = new JTextArea();
		textArea.setEditable(false);
		textArea.setBorder(new TitledBorder(null, "", TitledBorder.LEADING, TitledBorder.TOP, null, null));
		textArea.setLineWrap(true);
		textArea.setBounds(149, 51, 358, 83);
		contentPane.add(textArea);
		
		JLabel lblComment = new JLabel("Comment:");
		lblComment.setFont(new Font("Dialog", Font.BOLD, 12));
		lblComment.setBounds(58, 83, 73, 16);
		contentPane.add(lblComment);
		
		JCheckBox checkBox1 = new JCheckBox("Accessability of teacher outside classroom");
		checkBox1.setBounds(149, 162, 263, 24);
		contentPane.add(checkBox1);
		
		JCheckBox checkBox3 = new JCheckBox("Instructor's ability to motivate you towards module");
		checkBox3.setBounds(149, 216, 299, 24);
		contentPane.add(checkBox3);
		
		JCheckBox checkBox5 = new JCheckBox("Adherence to course outlines?");
		checkBox5.setBounds(149, 272, 196, 24);
		contentPane.add(checkBox5);
		
		JCheckBox checkBox6 = new JCheckBox("Instructor's concerns regarding labs");
		checkBox6.setBounds(149, 299, 226, 24);
		contentPane.add(checkBox6);
		
		JCheckBox checkBox2 = new JCheckBox("Knowledge Base/grip of instructor over subject");
		checkBox2.setBounds(149, 189, 283, 24);
		contentPane.add(checkBox2);
		
		JCheckBox checkBox4 = new JCheckBox("Instructor's ability to integrate contents of module with real world");
		checkBox4.setBounds(149, 245, 377, 24);
		contentPane.add(checkBox4);
		
		JCheckBox checkBox7 = new JCheckBox("Your satisfaction level with delivery method of instructor");
		checkBox7.setBounds(149, 326, 329, 24);
		contentPane.add(checkBox7);
		
		JButton btnNext = new JButton("Next Comment");
		btnNext.setFont(new Font("Dialog", Font.BOLD, 12));
		btnNext.setBounds(149, 385, 128, 24);
		contentPane.add(btnNext);
		
		JButton btnSkip = new JButton("Skip");
		btnSkip.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				int[] list=new int[10];
				String str=textArea.getText();
				list[0]=sentences.indexOf(str);
				list[8]=1;
				updateSheet(list);
				str=getRandomSentence();
				textArea.setText(str);
				checkBox1.setSelected(false);
				checkBox2.setSelected(false);
				checkBox3.setSelected(false);
				checkBox4.setSelected(false);
				checkBox5.setSelected(false);
				checkBox6.setSelected(false);
				checkBox7.setSelected(false);
			}
		});
		btnSkip.setFont(new Font("Dialog", Font.BOLD, 12));
		btnSkip.setBounds(358, 385, 78, 24);
		contentPane.add(btnSkip);
		textArea.setText(getRandomSentence());
		btnNext.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				int[] list=new int[10];
				String str=textArea.getText();
				list[0]=sentences.indexOf(str);
				if(checkBox1.isSelected()){
					list[1]=1;
				}
				
				if(checkBox2.isSelected()){
					list[2]=1;
				}
				if(checkBox3.isSelected()){
					list[3]=1;
				}
				
				if(checkBox4.isSelected()){
					list[4]=1;
				}
				
				if(checkBox5.isSelected()){
					list[5]=1;
				}
				if(checkBox6.isSelected()){
					list[6]=1;
				}
				if(checkBox7.isSelected()){
					list[7]=1;
				}
				list[8]=0;
				list[9]=1;
				updateSheet(list);
				str=getRandomSentence();
				textArea.setText(str);
				checkBox1.setSelected(false);
				checkBox2.setSelected(false);
				checkBox3.setSelected(false);
				checkBox4.setSelected(false);
				checkBox5.setSelected(false);
				checkBox6.setSelected(false);
				checkBox7.setSelected(false);
			}
		});
	}
	public ArrayList<String> readSentences(){
		ArrayList<String> list=new ArrayList<>();
		try {
			FileInputStream file=new FileInputStream(new File("output.xlsx"));
			XSSFWorkbook workbook=new XSSFWorkbook(file);
			XSSFSheet sheet=workbook.getSheetAt(1);
			Iterator<Row> rowIterator=sheet.iterator();
			rowIterator.next();
			while(rowIterator.hasNext()){
				Row row=rowIterator.next();
				String str=row.getCell(2).getStringCellValue();
				list.add(str);
			}
			
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		return list;
	}
	private void updateSheet(int[] values){
		XSSFWorkbook workbook;
		try {
			File f=new File("output.xlsx");
			FileInputStream file=new FileInputStream(f);
			workbook = new XSSFWorkbook(file);
			XSSFSheet sheet=workbook.getSheetAt(1);
			Row row=sheet.getRow(values[0]+1);
			Cell c1=row.getCell(3);
			c1.setCellValue(c1.getNumericCellValue()+values[1]);
			Cell c2=row.getCell(4);
			c2.setCellValue(c2.getNumericCellValue()+values[2]);
			Cell c3=row.getCell(5);
			c3.setCellValue(c3.getNumericCellValue()+values[3]);
			Cell c4=row.getCell(6);
			c4.setCellValue(c4.getNumericCellValue()+values[4]);
			Cell c5=row.getCell(7);
			c5.setCellValue(c5.getNumericCellValue()+values[5]);
			Cell c6=row.getCell(8);
			c6.setCellValue(c6.getNumericCellValue()+values[6]);
			Cell c7=row.getCell(9);
			c7.setCellValue(c7.getNumericCellValue()+values[7]);
			Cell c8=row.getCell(10);
			c8.setCellValue(c8.getNumericCellValue()+values[8]);
			Cell c9=row.getCell(11);
			c9.setCellValue(c9.getNumericCellValue()+values[9]);
			file.close();
			FileOutputStream fos= new FileOutputStream(f);
			workbook.write(fos);
			fos.close();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	private String getRandomSentence(){
		//makes sure that no sentence is left....
		if(temp.size()==0){
			temp.addAll(sentences);
		}
		Random r=new Random();
		int index=r.nextInt(temp.size());
		String str=temp.get(index);
		temp.remove(index);
		return str;
	}
}
