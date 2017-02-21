import org.zeroturnaround.zip.ZipUtil
import java.awt.BorderLayout
import java.io.File
import javax.swing.*
import javax.swing.filechooser.FileNameExtensionFilter



fun main(args: Array<String>) {
    SwingUtilities.invokeLater {
        UIManager.put("swing.boldMetal", false)
        val frame = JFrame("Word Unlocker")
        frame.defaultCloseOperation = JFrame.EXIT_ON_CLOSE
        frame.add(WordUnlocker())
        frame.pack()
        frame.setLocationRelativeTo(null)
        frame.isVisible = true
    }
}

fun unlock(file : File) {
    val entryPath = "word/settings.xml"
    val settings = String(ZipUtil.unpackEntry(file, entryPath))
    ZipUtil.replaceEntry(file, entryPath,
            settings.replace("w:enforcement=\"1\"", "w:enforcement=\"0\"").toByteArray())

}

class WordUnlocker : JPanel {
    constructor() : super(BorderLayout()) {
        val log = JTextArea(3, 20)
        val fc = JFileChooser()
        val openButton = JButton("打开文件")

        fc.fileFilter = FileNameExtensionFilter(
                "Word docx documents", "docx")

        log.isEditable = false
        log.lineWrap = true
        log.text = "请选择docx文件..."
        openButton.addActionListener {
            val returnVal = fc.showOpenDialog(this)

            if (returnVal == JFileChooser.APPROVE_OPTION) {
                val file = fc.selectedFile
                unlock(file)
                log.text = ("「" + file.name + "」\n 文件解锁成功" + "\n")
            }
        }

        val leftPanel = JPanel(BorderLayout())
        leftPanel.add(openButton, BorderLayout.NORTH)
        leftPanel.add(JLabel("written by Kymair"))

        add(leftPanel, BorderLayout.WEST)
        add(log)
    }
}