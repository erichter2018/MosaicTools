using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MosaicTools.Services;

/// <summary>
/// Get Prior logic for extracting and formatting prior study information.
/// Matches Python's perform_get_prior() and related helper methods.
/// </summary>
public class GetPriorService
{
    /// <summary>
    /// Extract prior study info from clipboard text and format it.
    /// Returns the formatted COMPARISON string.
    /// </summary>
    public string? ProcessPriorText(string rawText)
    {
        if (string.IsNullOrWhiteSpace(rawText) || rawText.Length < 5)
            return null;
        
        try
        {
            // 1. Parse Date and Time
            string priorDate = "";
            string priorTimeFormatted = "";
            
            var dateMatch = Regex.Match(rawText, @"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec).*?(19\d{2}|20\d{2})", RegexOptions.IgnoreCase);
            var timeMatch = Regex.Match(rawText, @"(\d{1,2}:\d{2}:\d{2}\s+(?:E|C|M|P|AK?|H)[SD]T)", RegexOptions.IgnoreCase);
            
            if (dateMatch.Success && timeMatch.Success)
            {
                var monthMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
                {
                    ["JAN"] = 1, ["FEB"] = 2, ["MAR"] = 3, ["APR"] = 4, ["MAY"] = 5, ["JUN"] = 6,
                    ["JUL"] = 7, ["AUG"] = 8, ["SEP"] = 9, ["OCT"] = 10, ["NOV"] = 11, ["DEC"] = 12
                };
                
                var rawDateStr = dateMatch.Value;
                var year = dateMatch.Groups[2].Value;
                
                int monthNum = 0;
                foreach (var (m, n) in monthMap)
                {
                    if (rawDateStr.ToUpper().Contains(m))
                    {
                        monthNum = n;
                        break;
                    }
                }
                
                var dayMatch = Regex.Match(rawDateStr, @"(\d{1,2})\s+\d");
                string day = dayMatch.Success ? dayMatch.Groups[1].Value : "1";
                priorDate = $"{monthNum}/{int.Parse(day)}/{year}";
                
                // Format time (remove seconds)
                priorTimeFormatted = Regex.Replace(timeMatch.Groups[1].Value, @"(\d{1,2}:\d{2}):\d{2}", "$1");
            }
            
            // 2. Handle Status Flags
            var priorOriginal = rawText;
            foreach (var flag in new[] { "IN_PROGRESS", "NO_HL7_ORDER", "UNKNOWN", "SIGNED" })
            {
                priorOriginal = priorOriginal.Replace("\t" + flag, " SIGNXED");
            }
            
            string priorImages = "";
            if (priorOriginal.Contains("\tNO_IMAGES"))
            {
                priorImages = "No Prior Images. ";
            }
            
            // 3. Modality Specific Processing
            string priorDescript1 = "";
            
            if (priorOriginal.Contains("\tUS"))
            {
                priorDescript1 = ProcessUltrasound(priorOriginal);
            }
            else if (priorOriginal.Contains("\tMR"))
            {
                priorDescript1 = ProcessMR(priorOriginal);
            }
            else if (priorOriginal.Contains("\tNM"))
            {
                priorDescript1 = ProcessNM(priorOriginal);
            }
            else if (priorOriginal.Contains("\tXR") || priorOriginal.Contains("\tCR") || priorOriginal.Contains("\tX-ray"))
            {
                priorDescript1 = ProcessRadiograph(priorOriginal);
            }
            else if (priorOriginal.Contains("\tCT"))
            {
                priorDescript1 = ProcessCT(priorOriginal);
            }
            
            // 4. Final Formatting
            priorDescript1 = CapitalizeAcronyms(priorDescript1);
            
            bool includeTime = false;
            if (!string.IsNullOrEmpty(priorDate) && !string.IsNullOrEmpty(priorTimeFormatted))
            {
                if (DateTime.TryParseExact(priorDate, "M/d/yyyy", 
                    System.Globalization.CultureInfo.InvariantCulture,
                    System.Globalization.DateTimeStyles.None, out var pDt))
                {
                    var diffDays = (DateTime.Now - pDt).Days;
                    if (diffDays >= 0 && diffDays <= 2)
                        includeTime = true;
                }
            }
            
            var timeComponent = includeTime ? " " + priorTimeFormatted : "";
            var finalText = $" COMPARISON: {priorDate}{timeComponent} {priorDescript1}. {priorImages}";
            finalText = Regex.Replace(finalText, @" {2,}", " ").Replace(" .", ".").Trim();
            
            return finalText;
        }
        catch (Exception ex)
        {
            Logger.Trace($"Get Prior processing failed: {ex.Message}");
            return null;
        }
    }
    
    private string ProcessUltrasound(string text)
    {
        var match = Regex.Match(text, @"US.*SIGNXED", RegexOptions.IgnoreCase);
        if (!match.Success) return "";
        
        var desc = match.Value;
        desc = Regex.Replace(desc, @"^US", "", RegexOptions.IgnoreCase);
        desc = desc.Replace(" SIGNXED", "").Replace(" abd.", " abdomen.").Trim();
        desc = ReorderLaterality(desc.ToLower());
        
        if (desc.Contains(" with and without"))
            desc = Regex.Replace(desc, @"(\s+)(with and without)", " ultrasound$2", RegexOptions.IgnoreCase);
        else if (desc.Contains(" without"))
            desc = Regex.Replace(desc, @"(\s+)(without)", " ultrasound$2", RegexOptions.IgnoreCase);
        else if (desc.Contains(" with"))
            desc = Regex.Replace(desc, @"(\s+)(with)", " ultrasound$2", RegexOptions.IgnoreCase);
        else
            desc += " ultrasound";
        
        return desc;
    }
    
    private string ProcessMR(string text)
    {
        var match = Regex.Match(text, @"MR.*SIGNXED", RegexOptions.IgnoreCase);
        if (!match.Success) return "";
        
        var desc = match.Value;
        desc = Regex.Replace(desc, @"^MR", "", RegexOptions.IgnoreCase);
        desc = desc.Replace(" SIGNXED", "").Trim();
        desc = desc.Replace(" + ", " and ").Replace(" W/O", " without").Replace(" W/", " with");
        desc = desc.Replace(" W WO", " with and without").Replace(" WO", " without").Replace(" IV ", " ");
        
        desc = ReorderLaterality(desc.ToLower());
        
        bool modifierFound = false;
        if (desc.Contains(" mra") || desc.Contains(" mrv"))
        {
            modifierFound = true;
        }
        else if (desc.Contains(" angiography"))
        {
            desc = Regex.Replace(desc, @"\s+(angiography)", " MR $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (desc.Contains(" venography"))
        {
            desc = Regex.Replace(desc, @"\s+(venography)", " MR $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (desc.Contains(" with and without"))
        {
            desc = Regex.Replace(desc, @"\s+(with and without)", " MR $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (desc.Contains(" without"))
        {
            desc = Regex.Replace(desc, @"\s+(without)", " MR $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (desc.Contains(" with"))
        {
            desc = Regex.Replace(desc, @"\s+(with)", " MR $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        
        if (!modifierFound)
            desc += " MR";
        
        return desc;
    }
    
    private string ProcessNM(string text)
    {
        var match = Regex.Match(text, @"NM.*SIGNXED", RegexOptions.IgnoreCase);
        if (!match.Success) return "";
        
        var desc = match.Value.Replace("NM", "nuclear medicine").Replace(" SIGNXED", "").Trim().ToLower();
        return ReorderLaterality(desc);
    }
    
    private string ProcessRadiograph(string text)
    {
        Match match;
        if (text.Contains("\tXR"))
            match = Regex.Match(text, @"XR.*SIGNXED", RegexOptions.IgnoreCase);
        else if (text.Contains("\tCR"))
            match = Regex.Match(text, @"CR.*SIGNXED", RegexOptions.IgnoreCase);
        else
            match = Regex.Match(text, @"X-ray.*SIGNXED", RegexOptions.IgnoreCase);
        
        if (!match.Success) return "";
        
        var desc = match.Value;
        desc = Regex.Replace(desc, @"^(XR|CR|X-ray, OT|X-ray)\t?", "", RegexOptions.IgnoreCase);
        desc = desc.Replace(" SIGNXED", "").Trim();
        
        desc = ProcessRadiographDescription(desc.ToLower());
        return desc;
    }
    
    private string ProcessRadiographDescription(string text)
    {
        text = text.Replace(" vw", " view(s)").Replace(" 2v", " PA and lateral");
        text = text.Replace(" pa lat", " PA and lateral").Replace(" (kub)", "").Trim();
        
        text = ReorderLaterality(text);
        
        bool modifierFound = false;
        if (Regex.IsMatch(text, @"\d+\s+or\s+more\s+radiograph\s+views?", RegexOptions.IgnoreCase))
        {
            text = Regex.Replace(text, @"(\d+\s+or\s+more)\s+radiograph\s+(views?)", "radiograph $1 $2", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (Regex.IsMatch(text, @"\s\d+\s*view"))
        {
            text = Regex.Replace(text, @"(\s)(\d+\s*view)", "$1radiograph $2");
            modifierFound = true;
        }
        else if (text.Contains(" pa and lateral"))
        {
            text = text.Replace(" pa and lateral", " radiograph PA and lateral");
            modifierFound = true;
        }
        else if (text.Contains(" view"))
        {
            text = text.Replace(" view", " radiograph view");
            modifierFound = true;
        }
        
        if (!modifierFound)
            text += " radiograph";
        
        return text;
    }
    
    private string ProcessCT(string text)
    {
        var tempText = Regex.Replace(text, @"Oct", "OcX", RegexOptions.IgnoreCase);
        var match = Regex.Match(tempText, @"CT.*SIGNXED", RegexOptions.IgnoreCase);
        if (!match.Success) return "";
        
        var desc = match.Value;
        desc = desc.Replace("CTA - ", "CTA ");
        desc = desc.Replace("CTA", "CTAPLACEHOLDER");
        desc = Regex.Replace(desc, @"^CT(\s|$)", "$1", RegexOptions.IgnoreCase);
        desc = desc.Replace("CTAPLACEHOLDER", "CTA").Replace(" SIGNXED", "").Trim();
        
        // Substitutions
        desc = desc.Replace(" + ", " and ").Replace("+", " and ").Replace(" imags", "").Replace("Head Or Brain", "brain");
        desc = desc.Replace(" W/CONTRST INCL W/O", " with and without contrast").Replace(" W/O", " without");
        desc = desc.Replace(" W/ ", " with ").Replace(" W/", " with ");
        desc = Regex.Replace(desc, @"\s+W\s+", " with ", RegexOptions.IgnoreCase);
        desc = desc.Replace(" W WO", " with and without").Replace(" WO", " without").Replace(" IV ", " ");
        
        desc = desc.Replace("ab pe", "abdomen and pelvis").Replace("abd & pelvis", "abdomen and pelvis");
        desc = desc.Replace("abd/pelvis", "abdomen and pelvis").Replace(" abd pel ", " abdomen and pelvis ");
        desc = desc.Replace("abdomen/pelvis", "abdomen and pelvis").Replace("chest/abdomen/pelvis", "chest, abdomen, and pelvis");
        desc = desc.Replace("Thorax", "chest").Replace("thorax", "chest").Replace("P.E", "PE").Replace("p.e", "PE");
        desc = Regex.Replace(desc, @"\s+protocol\s*$", "", RegexOptions.IgnoreCase);
        
        desc = ReorderLaterality(desc.ToLower());
        
        bool modifierFound = false;
        bool hasCta = false;
        bool ctaMovedFromStart = false;
        
        if (Regex.IsMatch(desc, @"^cta\s+", RegexOptions.IgnoreCase))
        {
            desc = Regex.Replace(desc, @"^cta\s+", "", RegexOptions.IgnoreCase);
            hasCta = true;
            ctaMovedFromStart = true;
            modifierFound = true;
        }
        else if (desc.Contains(" cta"))
        {
            hasCta = true;
            modifierFound = true;
        }
        else if (desc.Contains(" angiography"))
        {
            desc = Regex.Replace(desc, @"\s+(angiography)", " CT $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (desc.Contains(" with and without"))
        {
            desc = Regex.Replace(desc, @"\s+(with and without)", " CT $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (desc.Contains(" without"))
        {
            desc = Regex.Replace(desc, @"\s+(without)", " CT $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        else if (desc.Contains(" with"))
        {
            desc = Regex.Replace(desc, @"\s+(with)", " CT $1", RegexOptions.IgnoreCase);
            modifierFound = true;
        }
        
        if (ctaMovedFromStart)
        {
            if (Regex.IsMatch(desc, @"^(right|left|bilateral)\s+", RegexOptions.IgnoreCase))
                desc = Regex.Replace(desc, @"^((?:right|left|bilateral)\s+\w+)\s*", "$1 CTA ", RegexOptions.IgnoreCase);
            else
                desc = Regex.Replace(desc, @"^(\w+)\s*", "$1 CTA ", RegexOptions.IgnoreCase);
        }
        
        if (!modifierFound && !hasCta)
            desc += " CT";
        
        return desc;
    }
    
    private static string ReorderLaterality(string text)
    {
        if (Regex.IsMatch(text, @"^(right|left|bilateral)\b", RegexOptions.IgnoreCase))
            return text;
        
        var match = Regex.Match(text, @"\b(right|left|bilateral)\b", RegexOptions.IgnoreCase);
        if (match.Success)
        {
            var term = match.Groups[1].Value;
            var cleaned = Regex.Replace(text, $@"\s*\b{term}\b\s*", " ", RegexOptions.IgnoreCase);
            return $"{term} {cleaned.Trim()}".Trim();
        }
        
        return text;
    }
    
    private static string CapitalizeAcronyms(string text)
    {
        text = Regex.Replace(text, @"\bcta\b", "CTA", RegexOptions.IgnoreCase);
        text = Regex.Replace(text, @"\bct\b", "CT", RegexOptions.IgnoreCase);
        
        foreach (var term in new[] { "MRA", "MRV", "MRI", "MR", "PA", "PE" })
        {
            text = Regex.Replace(text, $@"\b{term}\b", term, RegexOptions.IgnoreCase);
        }
        
        return text;
    }
}
