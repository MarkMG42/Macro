import java.io.*;
import java.nio.file.*;
import org.apache.poi.poifs.filesystem.*;

/** Replace specified streams while retaining the original POIFS host storages. */
public class ReplaceOleStreams {
    private static void copy(DirectoryEntry from, DirectoryEntry to, Path replacements, String prefix) throws IOException {
        to.setStorageClsid(from.getStorageClsid());
        for (Entry entry : from) {
            String name = entry.getName(), path = prefix + name;
            if (entry.isDirectoryEntry()) {
                copy((DirectoryEntry) entry, to.createDirectory(name), replacements, path + "/");
            } else {
                Path replacement = replacements.resolve(path);
                try (InputStream content = Files.isRegularFile(replacement) ? Files.newInputStream(replacement)
                        : new DocumentInputStream((DocumentEntry) entry)) {
                    to.createDocument(name, content);
                }
            }
        }
    }

    public static void main(String[] args) throws Exception {
        if (args.length != 3) throw new IllegalArgumentException("input.swp output.swp replacements-directory");
        Path input = Path.of(args[0]), output = Path.of(args[1]), base = Path.of(args[2]);
        if (input.toRealPath().equals(output.toAbsolutePath().normalize()))
            throw new IllegalArgumentException("Never overwrite the original macro");
        // A fresh filesystem prevents freed/slack sectors from retaining obsolete p-code.
        try (InputStream in = Files.newInputStream(input);
             POIFSFileSystem original = new POIFSFileSystem(in);
             POIFSFileSystem rebuilt = new POIFSFileSystem()) {
            copy(original.getRoot(), rebuilt.getRoot(), base, "");
            try (OutputStream out = Files.newOutputStream(output)) { rebuilt.writeFilesystem(out); }
        }
    }
}
