package se.alipsa.excelutils;

import java.io.File;
import java.net.MalformedURLException;
import java.net.URISyntaxException;
import java.net.URL;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;

/**
 * Common file utilities
 */
public class FileUtil {

   private FileUtil() {
      // prevent instantiation
   }

   /**
    * Verify that the filePath exists and is reachable
    * @param filePath the path + file name of the resource to find
    * @return a File of found
    * @throws Exception if the filePath cannot be found
    */
   public static File checkFilePath(String filePath) throws Exception {
      File excelFile;
      URL url = getResourceUrl(filePath);
      if (url == null) {
         throw new Exception(filePath + " does not exist");
      }
      try {
         excelFile = Paths.get(url.toURI()).toFile();
      } catch (URISyntaxException | RuntimeException e) {
         throw new Exception(filePath + " does not exist");
      }
      if (!excelFile.exists()) {
         throw new Exception(filePath + " does not exist");
      }
      return excelFile;
   }

   /**
    * Find a resource using available class loaders.
    * It will also load resources/files from the
    * absolute path of the file system (not only the classpath's).
    * @param resource the resource to search for
    * @return an url to the resource or null if not found
    */
   public static URL getResourceUrl(String resource) {
      final List<ClassLoader> classLoaders = new ArrayList<ClassLoader>();
      classLoaders.add(Thread.currentThread().getContextClassLoader());
      classLoaders.add(FileUtil.class.getClassLoader());

      for (ClassLoader classLoader : classLoaders) {
         final URL url = getResourceWith(classLoader, resource);
         if (url != null) {
            return url;
         }
      }

      final URL systemResource = ClassLoader.getSystemResource(resource);
      if (systemResource != null) {
         return systemResource;
      } else {
         try {
            return new File(resource).toURI().toURL();
         } catch (MalformedURLException e) {
            return null;
         }
      }
   }

   private static URL getResourceWith(ClassLoader classLoader, String resource) {
      if (classLoader != null) {
         return classLoader.getResource(resource);
      }
      return null;
   }
}
