using System.Collections.Concurrent;
using System.IO;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Flow.Launcher.Plugin.QuickTodo.Services;

/// <summary>
/// Renders result icons ourselves via <see cref="Result.Icon"/> instead of letting
/// Flow Launcher resolve <see cref="Result.IcoPath"/> through the Windows shell
/// thumbnail provider. On some machines that provider fails to instantiate
/// (REGDB_E_CLASSNOTREG) and FL substitutes its missing-image placeholder for every
/// plugin's small list-row icon, even though the file is valid. Decoding the PNG into
/// a frozen <see cref="BitmapImage"/> sidesteps that path entirely. Images are cached
/// by absolute path and frozen so the delegate is safe to invoke on the UI thread.
/// </summary>
public class IconLoader
{
    private readonly string _pluginDirectory;
    private readonly ConcurrentDictionary<string, ImageSource?> _cache =
        new(StringComparer.OrdinalIgnoreCase);

    public IconLoader(string pluginDirectory)
    {
        _pluginDirectory = pluginDirectory;
    }

    /// <summary>
    /// Sets <see cref="Result.Icon"/> on every result whose IcoPath resolves to a file
    /// that decodes successfully. Results with an unresolved/unreadable icon are left
    /// untouched so FL falls back to its normal IcoPath handling.
    /// </summary>
    public List<Result> Apply(List<Result> results)
    {
        foreach (var result in results)
        {
            var image = LoadIcon(result.IcoPath);
            if (image == null)
                continue;

            // FL only consults Result.Icon when IcoPath is empty; otherwise it resolves
            // IcoPath through the (broken-on-this-machine) shell thumbnail provider. Clear
            // IcoPath so our directly-decoded image is what actually renders.
            result.Icon = () => image;
            result.IcoPath = string.Empty;
        }
        return results;
    }

    /// <summary>
    /// Resolves a plugin-relative (or absolute) IcoPath and decodes it into a frozen,
    /// cached image, or returns null if it can't be resolved/decoded. Used to swap a
    /// row's icon at runtime (e.g. toggling a task done) via <see cref="Result.Icon"/>.
    /// </summary>
    public ImageSource? LoadIcon(string? icoPath)
    {
        var path = Resolve(icoPath);
        return path == null ? null : Load(path);
    }

    private string? Resolve(string? icoPath)
    {
        if (string.IsNullOrWhiteSpace(icoPath))
            return null;

        var full = Path.IsPathRooted(icoPath)
            ? icoPath
            : Path.Combine(_pluginDirectory, icoPath);
        return File.Exists(full) ? full : null;
    }

    private ImageSource? Load(string fullPath)
    {
        return _cache.GetOrAdd(fullPath, p =>
        {
            try
            {
                var bmp = new BitmapImage();
                bmp.BeginInit();
                bmp.CacheOption = BitmapCacheOption.OnLoad;
                bmp.CreateOptions = BitmapCreateOptions.IgnoreImageCache;
                bmp.UriSource = new Uri(p, UriKind.Absolute);
                bmp.EndInit();
                bmp.Freeze();
                return bmp;
            }
            catch
            {
                return null;
            }
        });
    }
}
