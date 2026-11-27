package expo.modules.konvertpdf

import android.content.Context
import android.graphics.Bitmap
import android.net.Uri
import com.tom_roush.pdfbox.android.PDFBoxResourceLoader
import com.tom_roush.pdfbox.cos.COSBase
import com.tom_roush.pdfbox.cos.COSName
import com.tom_roush.pdfbox.io.MemoryUsageSetting
import com.tom_roush.pdfbox.multipdf.PDFMergerUtility
import com.tom_roush.pdfbox.pdmodel.PDDocument
import com.tom_roush.pdfbox.pdmodel.PDResources
import com.tom_roush.pdfbox.pdmodel.graphics.form.PDFormXObject
import com.tom_roush.pdfbox.pdmodel.graphics.image.JPEGFactory
import com.tom_roush.pdfbox.pdmodel.graphics.image.PDImageXObject
import expo.modules.kotlin.modules.Module
import expo.modules.kotlin.modules.ModuleDefinition
import java.io.ByteArrayInputStream
import java.io.ByteArrayOutputStream
import java.io.File
import java.io.FileInputStream
import java.util.Collections
import java.util.IdentityHashMap

class KonvertPdfModule : Module() {
  private var initialized = false

  private fun ctx(): Context {
    val c = appContext.reactContext ?: throw IllegalStateException("No Android context")
    if (!initialized) {
      PDFBoxResourceLoader.init(c.applicationContext)
      initialized = true
    }
    return c
  }

  private fun clean(p: String): String =
    if (p.startsWith("file://")) (Uri.parse(p).path ?: p.removePrefix("file://")) else p

  private fun openDoc(path: String): PDDocument =
    PDDocument.load(File(clean(path)), MemoryUsageSetting.setupTempFileOnly())

  override fun definition() = ModuleDefinition {
    Name("KonvertPdf")

    AsyncFunction("getPageCount") { path: String ->
      ctx()
      openDoc(path).use { it.numberOfPages }
    }

    AsyncFunction("mergePdfs") { paths: List<String>, outPath: String ->
      ctx()
      val out = File(clean(outPath))
      out.parentFile?.mkdirs()
      val merger = PDFMergerUtility()
      paths.forEach { merger.addSource(File(clean(it))) }
      merger.destinationFileName = out.absolutePath
      merger.mergeDocuments(MemoryUsageSetting.setupTempFileOnly())
      mapOf<String, Any>("path" to out.absolutePath, "outputBytes" to out.length().toDouble())
    }

    AsyncFunction("splitPdf") { path: String, pages: List<Int>, outPath: String ->
      ctx()
      val out = File(clean(outPath))
      out.parentFile?.mkdirs()
      openDoc(path).use { src ->
        PDDocument().use { dst ->
          for (i in pages) {
            dst.importPage(src.getPage(i))
          }
          dst.save(out)
        }
      }
      mapOf<String, Any>("path" to out.absolutePath, "outputBytes" to out.length().toDouble())
    }

    AsyncFunction("compressPdf") { path: String, outPath: String, maxDim: Int, quality: Int ->
      ctx()
      val out = File(clean(outPath))
      out.parentFile?.mkdirs()
      val stats = intArrayOf(0, 0) // replaced, total
      openDoc(path).use { doc ->
        val cache = IdentityHashMap<COSBase, PDImageXObject?>()
        val seen: MutableSet<COSBase> = Collections.newSetFromMap(IdentityHashMap<COSBase, Boolean>())
        for (page in doc.pages) {
          val res = page.resources
          if (res != null) {
            processResources(doc, res, cache, seen, maxDim, quality, stats)
          }
        }
        doc.save(out)
      }
      mapOf<String, Any>(
        "path" to out.absolutePath,
        "outputBytes" to out.length().toDouble(),
        "replaced" to stats[0].toDouble(),
        "total" to stats[1].toDouble()
      )
    }

    AsyncFunction("copyToUri") { srcPath: String, destUri: String ->
      val c = ctx()
      FileInputStream(File(clean(srcPath))).use { input ->
        val output = c.contentResolver.openOutputStream(Uri.parse(destUri), "w")
          ?: throw IllegalStateException("Cannot open destination")
        output.use { input.copyTo(it, 1024 * 1024) }
      }
      true
    }
  }

  private fun processResources(
    doc: PDDocument,
    res: PDResources,
    cache: IdentityHashMap<COSBase, PDImageXObject?>,
    seen: MutableSet<COSBase>,
    maxDim: Int,
    quality: Int,
    stats: IntArray
  ) {
    if (!seen.add(res.cosObject)) return
    for (name in res.xObjectNames.toList()) {
      val xo = try { res.getXObject(name) } catch (e: Throwable) { null }
      if (xo == null) continue
      if (xo is PDImageXObject) {
        stats[1]++
        val key: COSBase = xo.cosObject
        val rep: PDImageXObject? = if (cache.containsKey(key)) {
          cache[key]
        } else {
          val r = recompress(doc, xo, maxDim, quality)
          cache[key] = r
          if (r != null) cache[r.cosObject] = null
          r
        }
        if (rep != null) {
          res.put(name, rep)
          stats[0]++
        }
      } else if (xo is PDFormXObject) {
        val fr = xo.resources
        if (fr != null) {
          processResources(doc, fr, cache, seen, maxDim, quality, stats)
        }
      }
    }
  }

  private fun recompress(doc: PDDocument, img: PDImageXObject, maxDim: Int, quality: Int): PDImageXObject? {
    var bmp: Bitmap? = null
    var scaled: Bitmap? = null
    try {
      if (img.isStencil) return null
      val d = img.cosObject
      if (d.containsKey(COSName.SMASK) || d.containsKey(COSName.MASK)) return null
      if (img.bitsPerComponent == 1) return null
      if (img.colorSpace.numberOfComponents == 4) return null
      if (img.width < 150 && img.height < 150) return null
      val origSize = d.getInt(COSName.LENGTH, 0)
      val decoded = img.image ?: return null
      bmp = decoded
      val longest = maxOf(decoded.width, decoded.height)
      val target: Bitmap = if (longest > maxDim) {
        val s = maxDim.toFloat() / longest
        Bitmap.createScaledBitmap(
          decoded,
          maxOf(1, Math.round(decoded.width * s)),
          maxOf(1, Math.round(decoded.height * s)),
          true
        )
      } else {
        decoded
      }
      scaled = target
      val bos = ByteArrayOutputStream()
      target.compress(Bitmap.CompressFormat.JPEG, quality, bos)
      val bytes = bos.toByteArray()
      val downscaled = target !== decoded
      if (origSize > 0 && bytes.size >= origSize * 0.95) return null
      if (origSize <= 0 && !downscaled) return null
      return JPEGFactory.createFromStream(doc, ByteArrayInputStream(bytes))
    } catch (e: Throwable) {
      return null
    } finally {
      try {
        if (scaled != null && scaled !== bmp) scaled.recycle()
        bmp?.recycle()
      } catch (e: Throwable) {
      }
    }
  }
}
