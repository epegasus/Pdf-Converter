package dev.epegasus.pdfconverter

import android.Manifest.permission.READ_EXTERNAL_STORAGE
import android.Manifest.permission.WRITE_EXTERNAL_STORAGE
import android.content.Intent
import android.content.pm.PackageManager
import android.net.Uri
import android.os.Bundle
import android.os.Environment
import android.util.Log
import android.widget.Toast
import androidx.activity.result.contract.ActivityResultContracts
import androidx.appcompat.app.AppCompatActivity
import androidx.core.app.ActivityCompat
import androidx.core.content.ContextCompat
import com.aspose.words.WatermarkType
import com.itextpdf.io.image.ImageDataFactory
import com.itextpdf.kernel.geom.PageSize
import com.itextpdf.kernel.pdf.PdfDocument
import com.itextpdf.kernel.pdf.PdfWriter
import com.itextpdf.layout.Document
import com.itextpdf.layout.element.Image
import dev.epegasus.pdfconverter.databinding.ActivityMainBinding
import java.io.File
import java.io.FileNotFoundException
import java.io.IOException

private const val PERMISSION_REQUEST_CODE = 200
private const val TAG = "MyTag"

class MainActivity : AppCompatActivity() {

    private val binding by lazy { ActivityMainBinding.inflate(layoutInflater) }

    override fun onCreate(savedInstanceState: Bundle?) {
        super.onCreate(savedInstanceState)
        setContentView(binding.root)

        // apply the license if you have the Aspose.Words license...
        applyLicense()

        checkStoragePermission()

        binding.btnGenerateImagePdfMain.setOnClickListener { openGallery(0) }
        binding.btnGenerateWordPdfMain.setOnClickListener { openGallery(1) }
    }

    private fun applyLicense() {
        // set license
        /*val lic = License()
        val inputStream = resources.openRawResource(R.raw.license)
        try {
            lic.setLicense(inputStream);
        } catch (e: java.lang.Exception) {
            e.printStackTrace()
        }*/
    }

    private fun checkStoragePermission() {
        if (checkPermission()) {
            Toast.makeText(this, "Permission Granted", Toast.LENGTH_SHORT).show()
        } else {
            requestPermission()
        }
    }

    private fun checkPermission(): Boolean {
        val permission1 = ContextCompat.checkSelfPermission(applicationContext, WRITE_EXTERNAL_STORAGE)
        val permission2 = ContextCompat.checkSelfPermission(applicationContext, READ_EXTERNAL_STORAGE)
        return permission1 == PackageManager.PERMISSION_GRANTED && permission2 == PackageManager.PERMISSION_GRANTED
    }

    private fun requestPermission() {
        ActivityCompat.requestPermissions(this, arrayOf(WRITE_EXTERNAL_STORAGE, READ_EXTERNAL_STORAGE), PERMISSION_REQUEST_CODE)
    }

    override fun onRequestPermissionsResult(requestCode: Int, permissions: Array<out String>, grantResults: IntArray) {
        super.onRequestPermissionsResult(requestCode, permissions, grantResults)
        if (requestCode == PERMISSION_REQUEST_CODE) {
            if (grantResults.isNotEmpty()) {
                // after requesting permissions we are showing
                // users a toast message of permission granted.
                val writeStorage = grantResults[0] == PackageManager.PERMISSION_GRANTED
                val readStorage = grantResults[1] == PackageManager.PERMISSION_GRANTED
                if (writeStorage && readStorage) {
                    Toast.makeText(this, "Permission Granted..", Toast.LENGTH_SHORT).show()
                } else {
                    Toast.makeText(this, "Permission Denied.", Toast.LENGTH_SHORT).show()
                    finish()
                }
            }
        }
    }

    private val imagesResultLauncher = registerForActivityResult(ActivityResultContracts.GetMultipleContents()) {
        if (it.isEmpty()) {
            Toast.makeText(this, "No Images Selected", Toast.LENGTH_SHORT).show()
            return@registerForActivityResult
        }
        generateImagesPdf(it)
    }

    private val docResultLauncher = registerForActivityResult(ActivityResultContracts.StartActivityForResult()) {
        it.data?.let { intent ->
            intent.data?.let { uri ->
                generateWordPdf(uri)
                return@registerForActivityResult
            }
        }
        Toast.makeText(this, "No document Selected", Toast.LENGTH_SHORT).show()
    }

    /**
     * @param caseType
     *      0:  Images
     *      1:  Word
     */

    private fun openGallery(caseType: Int) {
        when (caseType) {
            0 -> imagesResultLauncher.launch("image/*")
            1 -> {
                val intent = Intent(Intent.ACTION_OPEN_DOCUMENT)
                intent.addCategory(Intent.CATEGORY_OPENABLE)
                // mime types for MS Word documents
                val mimetypes = arrayOf("application/vnd.openxmlformats-officedocument.wordprocessingml.document", "application/msword")
                intent.type = "*/*"
                intent.putExtra(Intent.EXTRA_MIME_TYPES, mimetypes)
                docResultLauncher.launch(intent)
            }
            2 -> {
                val pdfFolder = File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS), "Pdf Converter Testing")
                if (!pdfFolder.exists()) {
                    pdfFolder.mkdir()
                    Log.d(TAG, "generatePdf: folder created")
                }
                val doc = com.aspose.words.Document("$pdfFolder/PCW_1666628351793.pdf")
                if (doc.watermark.type == WatermarkType.TEXT) {
                    doc.watermark.remove()
                }
                doc.save("$pdfFolder/RemoveWatermark_out.pdf")
            }
        }
    }

    private fun generateImagesPdf(uris: List<Uri>) {

        // Path Directory
        val pdfFolder = File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS), "Pdf Converter Testing")
        if (!pdfFolder.exists()) {
            pdfFolder.mkdir()
            Log.d(TAG, "generatePdf: folder created")
        }

        // File Name
        val fileName = "PC_${System.currentTimeMillis()}.pdf"
        val pdfFile = File("$pdfFolder/$fileName")

        // Creating Document
        val pdfDocument = PdfDocument(PdfWriter(pdfFile))
        val document = Document(pdfDocument)
        val marginHorizontal = 60f
        val marginVertical = 50f

        document.setMargins(marginVertical, marginHorizontal, marginVertical, marginHorizontal)
        document.pdfDocument.defaultPageSize = PageSize.DEFAULT
        with(document.pdfDocument.documentInfo) {
            addCreationDate()
            title = fileName
            author = "Pdf Converter"
            creator = "Sohaib Ahmed"
        }

        val imageWidth = document.pdfDocument.defaultPageSize.width - (marginHorizontal + marginHorizontal)
        val imageHeight = document.pdfDocument.defaultPageSize.height - (marginVertical + marginVertical)

        // Adding Images
        uris.forEach {
            val imageData = ImageDataFactory.create(PathUtil.getPath(this, it))
            val image = Image(imageData)
            image.scaleToFit(imageWidth, imageHeight)

            /*
            // Align center
            image.setRelativePosition(
                (document.pdfDocument.defaultPageSize.width - image.imageScaledWidth) / 2,
                (document.pdfDocument.defaultPageSize.height - image.imageScaledHeight) / 2,
                (document.pdfDocument.defaultPageSize.width - image.imageScaledWidth) / 2,
                (document.pdfDocument.defaultPageSize.height - image.imageScaledHeight) / 2
            )*/
            document.add(image)
        }
        document.close()
    }

    private fun generateWordPdf(docUri: Uri) {
        // For complete examples and data files, please go to https://github.com/aspose-words/Aspose.Words-for-Java
        // Load the document from disk.
        // Path Directory
        val pdfFolder = File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOCUMENTS), "Pdf Converter Testing")
        if (!pdfFolder.exists()) {
            pdfFolder.mkdir()
            Log.d(TAG, "generatePdf: folder created")
        }

        // File Name
        val fileName = "PCW_${System.currentTimeMillis()}.pdf"
        val pdfFile = File("$pdfFolder/$fileName")

        // open the selected document into an Input stream
        try {
            contentResolver.openInputStream(docUri).use { inputStream ->
                val doc = com.aspose.words.Document(inputStream)
                // save DOCX as PDF
                doc.save(pdfFile.toString())
                // show PDF file location in toast as well as tree view (optional)
                Toast.makeText(this@MainActivity, "File saved in: $pdfFile", Toast.LENGTH_LONG).show()
            }
        } catch (e: FileNotFoundException) {
            e.printStackTrace()
            Toast.makeText(this@MainActivity, "File not found: " + e.message, Toast.LENGTH_LONG).show()
        } catch (e: IOException) {
            e.printStackTrace()
            Toast.makeText(this@MainActivity, e.message, Toast.LENGTH_LONG).show()
        } catch (e: Exception) {
            e.printStackTrace()
            Toast.makeText(this@MainActivity, e.message, Toast.LENGTH_LONG).show()
        }

    }
}