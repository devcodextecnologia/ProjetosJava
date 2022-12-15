package br.com.bulla.LeitorPlanilha;

import java.awt.Button;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.Image;
import java.awt.Toolkit;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.IOException;
import java.net.URL;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextArea;
import javax.swing.JTextPane;
import javax.swing.SwingConstants;

import com.jgoodies.forms.layout.ColumnSpec;
import com.jgoodies.forms.layout.FormLayout;
import com.jgoodies.forms.layout.FormSpecs;
import com.jgoodies.forms.layout.RowSpec;

public class TelaInicial {

	private JFrame frame;

	/**
	 * Launch the application.
	 */

	public static void main(String[] args) {

		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					TelaInicial window = new TelaInicial();
					window.frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public TelaInicial() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frame = new JFrame();
		frame.getContentPane().setFont(new Font("Tahoma", Font.PLAIN, 16));
		frame.setBounds(100, 100, 447, 454);
		frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		frame.getContentPane().setLayout(new FormLayout(
				new ColumnSpec[] { ColumnSpec.decode("436px:grow"), FormSpecs.RELATED_GAP_COLSPEC,
						FormSpecs.DEFAULT_COLSPEC, FormSpecs.RELATED_GAP_COLSPEC, FormSpecs.DEFAULT_COLSPEC,
						FormSpecs.RELATED_GAP_COLSPEC, FormSpecs.DEFAULT_COLSPEC, FormSpecs.RELATED_GAP_COLSPEC,
						FormSpecs.DEFAULT_COLSPEC, FormSpecs.RELATED_GAP_COLSPEC, FormSpecs.DEFAULT_COLSPEC, },
				new RowSpec[] { RowSpec.decode("22px"), FormSpecs.RELATED_GAP_ROWSPEC, FormSpecs.DEFAULT_ROWSPEC,
						FormSpecs.RELATED_GAP_ROWSPEC, FormSpecs.DEFAULT_ROWSPEC, FormSpecs.RELATED_GAP_ROWSPEC,
						FormSpecs.DEFAULT_ROWSPEC, FormSpecs.RELATED_GAP_ROWSPEC, FormSpecs.DEFAULT_ROWSPEC,
						FormSpecs.RELATED_GAP_ROWSPEC, FormSpecs.DEFAULT_ROWSPEC, FormSpecs.RELATED_GAP_ROWSPEC,
						RowSpec.decode("default:grow"), FormSpecs.RELATED_GAP_ROWSPEC, FormSpecs.DEFAULT_ROWSPEC, }));

		JTextArea txtrAasdsadasdas = new JTextArea();
		txtrAasdsadasdas.setToolTipText("");
		txtrAasdsadasdas.setFont(new Font("Monospaced", Font.PLAIN, 13));
		txtrAasdsadasdas.setColumns(2);
		txtrAasdsadasdas.setWrapStyleWord(true);
		txtrAasdsadasdas.setEditable(false);
		txtrAasdsadasdas.setTabSize(15);
		txtrAasdsadasdas.setText("       SISTEMA DE TRATAMENTO DE PLANILHAS V.01");
		frame.getContentPane().add(txtrAasdsadasdas, "1, 1, 3, 1, fill, top");

		JLabel lblNewLabel = new JLabel("");
		lblNewLabel.setVerticalAlignment(SwingConstants.TOP);
		lblNewLabel.setIcon(new ImageIcon(TelaInicial.class.getResource("/br/com/bulla/LeitorPlanilha/Bullla.png")));
		lblNewLabel.setHorizontalAlignment(SwingConstants.CENTER);
		frame.getContentPane().add(lblNewLabel, "1, 5");

		Button button = new Button("PROCESSAR PLANILHA");
		button.addActionListener(new ActionListener() {

			// clique no botao PROCESSAR
			public void actionPerformed(ActionEvent e) {

				// data/hora atual
				LocalDateTime agora = LocalDateTime.now();

				// formatar a hora e coloca na variavel hora como String
				DateTimeFormatter formatterHora = DateTimeFormatter.ofPattern("HHmmss");
				String hora = formatterHora.format(agora);

				Arquivo arquivo = new Arquivo();

				try {
					arquivo.criarListaDados(Dados.getDados(), "C:\\LerPlanilha\\PlanilhaTratada" + hora + ".xlsx");

					JOptionPane.showMessageDialog(null,
							"Arquivo Gerado Com Sucesso!!! \"C:\\\\LerPlanilha\\\\PlanilhaTratada\"" + hora
									+ "\".xlsx\"");

				} catch (IOException e1) {
					JOptionPane.showMessageDialog(null,
							"FALHA NO FUNCIONAMENTO!! \n 1-Verificar se existe a pasta c:\\LerPlanilha \n 2-Dentro da pasta deve colocar a planilha com o nome planilha.xlsx");
					e1.printStackTrace();
				}
			}
		});
		button.setFont(new Font("Dialog", Font.PLAIN, 8));
		frame.getContentPane().add(button, "1, 11, center, center");

		JTextPane txtpnEquipeDeSustentao = new JTextPane();
		txtpnEquipeDeSustentao.setText("                                       Equipe de Sustentação ao Negócio");
		frame.getContentPane().add(txtpnEquipeDeSustentao, "1, 15, fill, fill");
	}

}
