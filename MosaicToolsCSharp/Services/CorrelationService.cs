using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace MosaicTools.Services;

/// <summary>
/// A single correlation between an impression item and its matched findings.
/// </summary>
public class CorrelationItem
{
    public int ImpressionIndex { get; set; }
    public string ImpressionText { get; set; } = "";
    public List<string> MatchedFindings { get; set; } = new();
    public int ColorIndex { get; set; }
    public Color? HighlightColor { get; set; }
}

/// <summary>
/// Result of a correlation analysis between FINDINGS and IMPRESSION sections.
/// </summary>
public class CorrelationResult
{
    public List<CorrelationItem> Items { get; set; } = new();
    public string Source { get; set; } = "heuristic";
}

/// <summary>
/// Body-part categories for stable rainbow-mode coloring.
/// </summary>
public enum BodyPartCategory
{
    Brain, OrbitsSinusesEars, FaceSkull, NeckThroat, Thyroid, LymphNodes,
    Vascular, Heart, Mediastinum, LungsPleura, Airway,
    LiverBiliary, Spleen, Pancreas, KidneysAdrenals, GIBowel,
    PelvisReproductive, Peritoneum, SpineDiscs, BonesJoints, SoftTissues,
    LinesTubes, Other
}

/// <summary>
/// Heuristic-based correlation between FINDINGS and IMPRESSION sections of a radiology report.
/// Extracts canonical medical terms and matches impression items to their supporting findings.
/// </summary>
public static class CorrelationService
{
    /// <summary>
    /// Color palette for correlation highlighting (legacy fallback).
    /// </summary>
    public static readonly Color[] Palette = new[]
    {
        Color.FromArgb(255, 50, 50),    // Red
        Color.FromArgb(255, 140, 0),    // Orange
        Color.FromArgb(255, 255, 0),    // Yellow
        Color.FromArgb(0, 200, 0),      // Green
        Color.FromArgb(0, 150, 255),    // Blue
        Color.FromArgb(75, 0, 180),     // Indigo
        Color.FromArgb(180, 0, 255),    // Violet
    };

    /// <summary>
    /// Semantically chosen color for each body-part category.
    /// </summary>
    public static readonly Dictionary<BodyPartCategory, Color> CategoryColors = new()
    {
        { BodyPartCategory.Brain,              Color.FromArgb(180, 130, 255) }, // Purple
        { BodyPartCategory.OrbitsSinusesEars,   Color.FromArgb(200, 160, 255) }, // Light purple
        { BodyPartCategory.FaceSkull,           Color.FromArgb(170, 170, 220) }, // Slate blue
        { BodyPartCategory.NeckThroat,          Color.FromArgb(0, 200, 200) },   // Teal
        { BodyPartCategory.Thyroid,             Color.FromArgb(0, 180, 180) },   // Dark teal
        { BodyPartCategory.LymphNodes,          Color.FromArgb(180, 220, 100) }, // Yellow-green
        { BodyPartCategory.Vascular,            Color.FromArgb(255, 80, 80) },   // Bright red
        { BodyPartCategory.Heart,               Color.FromArgb(255, 50, 50) },   // Red
        { BodyPartCategory.Mediastinum,         Color.FromArgb(255, 140, 100) }, // Salmon
        { BodyPartCategory.LungsPleura,         Color.FromArgb(80, 160, 255) },  // Blue
        { BodyPartCategory.Airway,              Color.FromArgb(100, 200, 255) }, // Light blue
        { BodyPartCategory.LiverBiliary,        Color.FromArgb(180, 120, 60) },  // Brown
        { BodyPartCategory.Spleen,              Color.FromArgb(200, 80, 150) },  // Magenta-pink
        { BodyPartCategory.Pancreas,            Color.FromArgb(220, 180, 80) },  // Gold
        { BodyPartCategory.KidneysAdrenals,     Color.FromArgb(255, 180, 0) },   // Orange
        { BodyPartCategory.GIBowel,             Color.FromArgb(200, 200, 100) }, // Yellow-olive
        { BodyPartCategory.PelvisReproductive,  Color.FromArgb(255, 130, 180) }, // Pink
        { BodyPartCategory.Peritoneum,          Color.FromArgb(160, 180, 140) }, // Sage green
        { BodyPartCategory.SpineDiscs,          Color.FromArgb(100, 220, 100) }, // Green
        { BodyPartCategory.BonesJoints,         Color.FromArgb(200, 200, 200) }, // Light gray
        { BodyPartCategory.SoftTissues,         Color.FromArgb(180, 160, 140) }, // Tan
        { BodyPartCategory.LinesTubes,          Color.FromArgb(255, 255, 100) }, // Bright yellow
        { BodyPartCategory.Other,               Color.FromArgb(150, 150, 180) }, // Muted blue-gray
    };

    /// <summary>
    /// Maps ALL-CAPS subsection headers to body-part categories.
    /// </summary>
    private static readonly Dictionary<string, BodyPartCategory> HeaderToCategory =
        new(StringComparer.OrdinalIgnoreCase)
    {
        // Brain
        { "BRAIN", BodyPartCategory.Brain },
        { "BRAIN AND VENTRICLES", BodyPartCategory.Brain },
        { "BRAIN, ORBITS AND SINUSES", BodyPartCategory.Brain },
        { "CTA HEAD", BodyPartCategory.Brain },

        // Orbits, Sinuses, Ears
        { "ORBITS", BodyPartCategory.OrbitsSinusesEars },
        { "SINUSES", BodyPartCategory.OrbitsSinusesEars },
        { "SINUSES AND MASTOIDS", BodyPartCategory.OrbitsSinusesEars },
        { "MASTOID AIR CELLS", BodyPartCategory.OrbitsSinusesEars },
        { "EXTERNAL AUDITORY CANAL", BodyPartCategory.OrbitsSinusesEars },
        { "INNER EAR", BodyPartCategory.OrbitsSinusesEars },
        { "INTERNAL AUDITORY CANAL", BodyPartCategory.OrbitsSinusesEars },
        { "MIDDLE EAR CAVITY", BodyPartCategory.OrbitsSinusesEars },
        { "SALIVARY GLANDS", BodyPartCategory.OrbitsSinusesEars },

        // Face/Skull
        { "FACIAL BONES", BodyPartCategory.FaceSkull },
        { "SOFT TISSUES AND SKULL", BodyPartCategory.FaceSkull },

        // Neck/Throat
        { "AERODIGESTIVE TRACT", BodyPartCategory.NeckThroat },
        { "EPIGLOTTIS", BodyPartCategory.NeckThroat },
        { "CTA NECK", BodyPartCategory.NeckThroat },

        // Thyroid
        { "THYROID", BodyPartCategory.Thyroid },

        // Lymph Nodes
        { "LYMPH NODES", BodyPartCategory.LymphNodes },
        { "ABDOMINAL AND PELVIS LYMPH NODES", BodyPartCategory.LymphNodes },
        { "MEDIASTINUM AND LYMPH NODES", BodyPartCategory.LymphNodes },

        // Vascular
        { "AORTA", BodyPartCategory.Vascular },
        { "AORTIC ARCH AND ARCH VESSELS", BodyPartCategory.Vascular },
        { "CERVICAL CAROTID ARTERIES", BodyPartCategory.Vascular },
        { "CERVICAL VERTEBRAL ARTERIES", BodyPartCategory.Vascular },
        { "INTERNAL CAROTID ARTERIES", BodyPartCategory.Vascular },
        { "VERTEBRAL ARTERIES", BodyPartCategory.Vascular },
        { "ANTERIOR CEREBRAL ARTERIES", BodyPartCategory.Vascular },
        { "MIDDLE CEREBRAL ARTERIES", BodyPartCategory.Vascular },
        { "POSTERIOR CEREBRAL ARTERIES", BodyPartCategory.Vascular },
        { "ANTERIOR CIRCULATION", BodyPartCategory.Vascular },
        { "POSTERIOR CIRCULATION", BodyPartCategory.Vascular },
        { "BASILAR ARTERY", BodyPartCategory.Vascular },
        { "CELIAC TRUNK", BodyPartCategory.Vascular },
        { "SUPERIOR MESENTERIC ARTERY", BodyPartCategory.Vascular },
        { "INFERIOR MESENTERIC ARTERY", BodyPartCategory.Vascular },
        { "RENAL ARTERIES", BodyPartCategory.Vascular },
        { "ILIAC ARTERIES", BodyPartCategory.Vascular },
        { "LEFT ILIAC ARTERIES", BodyPartCategory.Vascular },
        { "RIGHT ILIAC ARTERIES", BodyPartCategory.Vascular },
        { "LEFT FEMORAL ARTERIES", BodyPartCategory.Vascular },
        { "RIGHT FEMORAL ARTERIES", BodyPartCategory.Vascular },
        { "RIGHT FEMORAL SRTERIES", BodyPartCategory.Vascular }, // typo in templates
        { "LEFT POPLITEAL ARTERY", BodyPartCategory.Vascular },
        { "RIGHT POPLITEAL ARTERY", BodyPartCategory.Vascular },
        { "LEFT CALF ARTERIES", BodyPartCategory.Vascular },
        { "RIGHT CALF ARTERIES", BodyPartCategory.Vascular },
        { "GREAT VESSELS OF AORTIC ARCH", BodyPartCategory.Vascular },
        { "PULMONARY ARTERIES", BodyPartCategory.Vascular },
        { "VASCULAR", BodyPartCategory.Vascular },
        { "VASCULATURE", BodyPartCategory.Vascular },
        { "VASULATURE", BodyPartCategory.Vascular }, // typo in templates

        // Heart
        { "CARDIAC MORPHOLOGY", BodyPartCategory.Heart },
        { "HEART AND MEDIASTINUM", BodyPartCategory.Heart },
        { "HEART AND PERICARDIUM", BodyPartCategory.Heart },
        { "RATE OF CARDIAC ACTIVITY", BodyPartCategory.Heart },

        // Mediastinum
        { "MEDIASTINUM", BodyPartCategory.Mediastinum },
        { "EXTRACARDIAC FINDINGS", BodyPartCategory.Mediastinum },

        // Lungs/Pleura
        { "LUNGS AND PLEURA", BodyPartCategory.LungsPleura },
        { "LUNGS AND MEDIASTINUM", BodyPartCategory.LungsPleura },
        { "CHEST", BodyPartCategory.LungsPleura },
        { "LOWER CHEST", BodyPartCategory.LungsPleura },
        { "LUNG BASE", BodyPartCategory.LungsPleura },

        // Liver/Biliary
        { "LIVER", BodyPartCategory.LiverBiliary },
        { "GALLBLADDER AND BILE DUCTS", BodyPartCategory.LiverBiliary },
        { "GALLBLADDER AND BILIARY SYSTEM", BodyPartCategory.LiverBiliary },
        { "HEPATOBILIARY", BodyPartCategory.LiverBiliary },

        // Spleen
        { "SPLEEN", BodyPartCategory.Spleen },

        // Pancreas
        { "PANCREAS", BodyPartCategory.Pancreas },
        { "PANCREAS/PANCREATIC DUCT", BodyPartCategory.Pancreas },

        // Kidneys/Adrenals
        { "KIDNEYS", BodyPartCategory.KidneysAdrenals },
        { "KIDNEYS, URETERS AND BLADDER", BodyPartCategory.KidneysAdrenals },
        { "ADRENAL GLANDS", BodyPartCategory.KidneysAdrenals },

        // GI/Bowel
        { "BOWEL", BodyPartCategory.GIBowel },
        { "GI AND BOWEL", BodyPartCategory.GIBowel },
        { "UPPER ABDOMEN", BodyPartCategory.GIBowel },
        { "ABDOMEN AND PELVIS", BodyPartCategory.GIBowel },

        // Pelvis/Reproductive
        { "REPRODUCTIVE", BodyPartCategory.PelvisReproductive },
        { "REPRODUCTIVE ORGANS", BodyPartCategory.PelvisReproductive },
        { "UTERUS", BodyPartCategory.PelvisReproductive },
        { "LEFT OVARY", BodyPartCategory.PelvisReproductive },
        { "RIGHT OVARY", BodyPartCategory.PelvisReproductive },
        { "INTRAPELVIC CONTENTS", BodyPartCategory.PelvisReproductive },
        { "GESTATIONAL SAC(S)", BodyPartCategory.PelvisReproductive },
        { "CROWN RUMP LENGTH", BodyPartCategory.PelvisReproductive },
        { "ESTIMATED DUE DATE", BodyPartCategory.PelvisReproductive },
        { "ESTIMATED GESTATIONAL AGE BY CURRENT ULTRASOUND", BodyPartCategory.PelvisReproductive },
        { "ESTIMATED GESTATIONAL AGE BY LMP/PRIOR ULTRASOUND", BodyPartCategory.PelvisReproductive },
        { "YOLK SAC", BodyPartCategory.PelvisReproductive },

        // Peritoneum
        { "PERITONEUM", BodyPartCategory.Peritoneum },
        { "PERITONEUM AND RETROPERITONEUM", BodyPartCategory.Peritoneum },
        { "PERITONEUM AND RETRPERITONEUM", BodyPartCategory.Peritoneum }, // typo
        { "FREE FLUID", BodyPartCategory.Peritoneum },
        { "ABDOMINAL WALL", BodyPartCategory.Peritoneum },

        // Spine/Discs
        { "LUMBAR SPINE", BodyPartCategory.SpineDiscs },
        { "THORACIC SPINE", BodyPartCategory.SpineDiscs },
        { "BONES AND ALIGNMENT", BodyPartCategory.SpineDiscs },
        { "DEGENERATIVE CHANGES", BodyPartCategory.SpineDiscs },
        { "DISC SPACES", BodyPartCategory.SpineDiscs },
        { "DISCS AND DEGENERATIVE CHANGES", BodyPartCategory.SpineDiscs },

        // Bones/Joints
        { "BONES", BodyPartCategory.BonesJoints },
        { "BONES AND JOINTS", BodyPartCategory.BonesJoints },
        { "BONE MARROW", BodyPartCategory.BonesJoints },
        { "JOINT SPACES", BodyPartCategory.BonesJoints },
        { "JOINTS", BodyPartCategory.BonesJoints },
        { "ABDOMINAL BONES AND SOFT TISSUES", BodyPartCategory.BonesJoints },
        { "THORACIC BONES AND SOFT TISSUES", BodyPartCategory.BonesJoints },

        // Soft Tissues
        { "SOFT TISSUES", BodyPartCategory.SoftTissues },
        { "SOFT TISSUES AND BONES", BodyPartCategory.SoftTissues },
        { "SOFT TISSUES/BONES", BodyPartCategory.SoftTissues },
        { "ACHILLES TENDON", BodyPartCategory.SoftTissues },
        { "ANTERIOR TENDONS", BodyPartCategory.SoftTissues },
        { "LATERAL TENDONS", BodyPartCategory.SoftTissues },
        { "MEDIAL TENDONS", BodyPartCategory.SoftTissues },
        { "DELTOID LIGAMENT COMPLEX", BodyPartCategory.SoftTissues },
        { "LATERAL COLLATERAL LIGAMENT COMPLEX", BodyPartCategory.SoftTissues },
        { "SYNDESMOTIC LIGAMENTS", BodyPartCategory.SoftTissues },
        { "SINUS TARSI AND SPRING LIGAMENT", BodyPartCategory.SoftTissues },
        { "TARSAL TUNNEL", BodyPartCategory.SoftTissues },
        { "PLANTAR FASCIA", BodyPartCategory.SoftTissues },

        // Lines/Tubes
        { "LINES, TUBES AND DEVICES", BodyPartCategory.LinesTubes },

        // Other
        { "OTHER", BodyPartCategory.Other },
        { "MISCELLANEOUS", BodyPartCategory.Other },
        { "EXAM QUALITY", BodyPartCategory.Other },
    };

    /// <summary>
    /// Keyword-based fallback for headers not in the explicit mapping.
    /// Checked in order; first match wins.
    /// </summary>
    private static readonly (string Keyword, BodyPartCategory Category)[] CategoryKeywords =
    {
        ("brain", BodyPartCategory.Brain),
        ("orbit", BodyPartCategory.OrbitsSinusesEars),
        ("sinus", BodyPartCategory.OrbitsSinusesEars),
        ("ear", BodyPartCategory.OrbitsSinusesEars),
        ("mastoid", BodyPartCategory.OrbitsSinusesEars),
        ("face", BodyPartCategory.FaceSkull),
        ("skull", BodyPartCategory.FaceSkull),
        ("neck", BodyPartCategory.NeckThroat),
        ("throat", BodyPartCategory.NeckThroat),
        ("thyroid", BodyPartCategory.Thyroid),
        ("lymph", BodyPartCategory.LymphNodes),
        ("arter", BodyPartCategory.Vascular),
        ("aort", BodyPartCategory.Vascular),
        ("vascular", BodyPartCategory.Vascular),
        ("vessel", BodyPartCategory.Vascular),
        ("cardiac", BodyPartCategory.Heart),
        ("heart", BodyPartCategory.Heart),
        ("pericardi", BodyPartCategory.Heart),
        ("mediastin", BodyPartCategory.Mediastinum),
        ("lung", BodyPartCategory.LungsPleura),
        ("pleura", BodyPartCategory.LungsPleura),
        ("chest", BodyPartCategory.LungsPleura),
        ("airway", BodyPartCategory.Airway),
        ("liver", BodyPartCategory.LiverBiliary),
        ("hepat", BodyPartCategory.LiverBiliary),
        ("gallbladder", BodyPartCategory.LiverBiliary),
        ("biliary", BodyPartCategory.LiverBiliary),
        ("spleen", BodyPartCategory.Spleen),
        ("pancrea", BodyPartCategory.Pancreas),
        ("kidney", BodyPartCategory.KidneysAdrenals),
        ("renal", BodyPartCategory.KidneysAdrenals),
        ("adrenal", BodyPartCategory.KidneysAdrenals),
        ("bowel", BodyPartCategory.GIBowel),
        ("abdomen", BodyPartCategory.GIBowel),
        ("pelvi", BodyPartCategory.PelvisReproductive),
        ("uterus", BodyPartCategory.PelvisReproductive),
        ("ovary", BodyPartCategory.PelvisReproductive),
        ("reproduct", BodyPartCategory.PelvisReproductive),
        ("peritoneum", BodyPartCategory.Peritoneum),
        ("retroperiton", BodyPartCategory.Peritoneum),
        ("spine", BodyPartCategory.SpineDiscs),
        ("disc", BodyPartCategory.SpineDiscs),
        ("disk", BodyPartCategory.SpineDiscs),
        ("degen", BodyPartCategory.SpineDiscs),
        ("bone", BodyPartCategory.BonesJoints),
        ("joint", BodyPartCategory.BonesJoints),
        ("soft tissue", BodyPartCategory.SoftTissues),
        ("tendon", BodyPartCategory.SoftTissues),
        ("ligament", BodyPartCategory.SoftTissues),
        ("line", BodyPartCategory.LinesTubes),
        ("tube", BodyPartCategory.LinesTubes),
        ("device", BodyPartCategory.LinesTubes),
    };

    /// <summary>
    /// Look up the body-part category for a subsection header.
    /// </summary>
    internal static BodyPartCategory GetCategoryForHeader(string? header)
    {
        if (string.IsNullOrWhiteSpace(header))
            return BodyPartCategory.Other;

        if (HeaderToCategory.TryGetValue(header.Trim(), out var cat))
            return cat;

        // Keyword fallback
        var lower = header.ToLowerInvariant();
        foreach (var (keyword, category) in CategoryKeywords)
        {
            if (lower.Contains(keyword))
                return category;
        }

        return BodyPartCategory.Other;
    }

    /// <summary>
    /// Get the stable color for a body-part category.
    /// </summary>
    internal static Color GetColorForCategory(BodyPartCategory category)
    {
        return CategoryColors.TryGetValue(category, out var color)
            ? color
            : CategoryColors[BodyPartCategory.Other];
    }

    /// <summary>
    /// Content-level keywords that map to a specific body-part category.
    /// Used to refine broad header categories (e.g., "MEDIASTINUM" containing heart findings).
    /// Checked in order; first match wins.
    /// </summary>
    private static readonly (string Keyword, BodyPartCategory Category)[] ContentKeywords =
    {
        ("heart", BodyPartCategory.Heart),
        ("pericardial", BodyPartCategory.Heart),
        ("pericardium", BodyPartCategory.Heart),
        ("cardiac", BodyPartCategory.Heart),
        ("cardiomegaly", BodyPartCategory.Heart),
        ("myocardial", BodyPartCategory.Heart),
        ("atrial", BodyPartCategory.Heart),
        ("ventricular", BodyPartCategory.Heart),
        ("valve", BodyPartCategory.Heart),
        ("pulmonary embol", BodyPartCategory.Vascular),
        ("aorta", BodyPartCategory.Vascular),
        ("aortic", BodyPartCategory.Vascular),
        ("aneurysm", BodyPartCategory.Vascular),
        ("liver", BodyPartCategory.LiverBiliary),
        ("hepatic", BodyPartCategory.LiverBiliary),
        ("gallbladder", BodyPartCategory.LiverBiliary),
        ("biliary", BodyPartCategory.LiverBiliary),
        ("spleen", BodyPartCategory.Spleen),
        ("splenic", BodyPartCategory.Spleen),
        ("pancrea", BodyPartCategory.Pancreas),
        ("kidney", BodyPartCategory.KidneysAdrenals),
        ("renal", BodyPartCategory.KidneysAdrenals),
        ("adrenal", BodyPartCategory.KidneysAdrenals),
        ("pleural", BodyPartCategory.LungsPleura),
        ("pulmonary", BodyPartCategory.LungsPleura),
        ("lung", BodyPartCategory.LungsPleura),
        ("nodule", BodyPartCategory.LungsPleura),
        ("thyroid", BodyPartCategory.Thyroid),
        ("lymph node", BodyPartCategory.LymphNodes),
        ("lymphadenopathy", BodyPartCategory.LymphNodes),
        ("fracture", BodyPartCategory.BonesJoints),
        ("osseous", BodyPartCategory.BonesJoints),
        ("skeletal", BodyPartCategory.BonesJoints),
        ("vertebra", BodyPartCategory.SpineDiscs),
        ("disc", BodyPartCategory.SpineDiscs),
        ("spine", BodyPartCategory.SpineDiscs),
    };

    /// <summary>
    /// Refine a header-based category by scanning the actual findings content.
    /// Only overrides when content points to a more specific category than the header.
    /// For example, "heart is enlarged" under MEDIASTINUM → Heart instead of Mediastinum.
    /// </summary>
    internal static BodyPartCategory RefineCategoryFromContent(BodyPartCategory headerCategory, List<string> findings)
    {
        // Only refine broad/container categories — specific ones are already correct
        if (headerCategory != BodyPartCategory.Mediastinum &&
            headerCategory != BodyPartCategory.GIBowel &&
            headerCategory != BodyPartCategory.SoftTissues &&
            headerCategory != BodyPartCategory.Other)
            return headerCategory;

        // Count keyword hits per category across all findings
        var hits = new Dictionary<BodyPartCategory, int>();
        var lowerFindings = string.Join(" ", findings).ToLowerInvariant();

        foreach (var (keyword, category) in ContentKeywords)
        {
            if (lowerFindings.Contains(keyword))
            {
                hits.TryGetValue(category, out int count);
                hits[category] = count + 1;
            }
        }

        if (hits.Count == 0)
            return headerCategory;

        // Return the category with the most keyword hits
        return hits.OrderByDescending(kv => kv.Value).First().Key;
    }

    /// <summary>
    /// Post-process correlation items so that items sharing the same HighlightColor
    /// get visually distinct variants (hue-rotated + lightness-shifted).
    /// Items with unique colors are left unchanged.
    /// </summary>
    internal static void DiversifyCategoryColors(List<CorrelationItem> items)
    {
        // Group by HighlightColor ARGB value
        var groups = new Dictionary<int, List<CorrelationItem>>();
        foreach (var item in items)
        {
            if (item.HighlightColor == null) continue;
            int key = item.HighlightColor.Value.ToArgb();
            if (!groups.ContainsKey(key))
                groups[key] = new List<CorrelationItem>();
            groups[key].Add(item);
        }

        foreach (var group in groups.Values)
        {
            if (group.Count <= 1) continue;

            var baseColor = group[0].HighlightColor!.Value;
            var (h, s, l) = RgbToHsl(baseColor);

            for (int i = 1; i < group.Count; i++)
            {
                int step = (i + 1) / 2;            // 1, 1, 2, 2, 3, 3, ...
                int sign = (i % 2 == 1) ? 1 : -1;  // +, -, +, -, ...

                double newH = h + sign * step * 25.0;
                double newL = l + sign * step * 0.07;

                // For near-neutral colors, inject saturation so hue rotation is visible
                double newS = s < 0.15 ? 0.20 + step * 0.05 : s;

                // Clamp
                newH = ((newH % 360.0) + 360.0) % 360.0;
                newS = Math.Clamp(newS, 0.0, 0.6);
                newL = Math.Clamp(newL, 0.25, 0.85);

                group[i].HighlightColor = HslToRgb(newH, newS, newL);
            }
        }
    }

    private static (double H, double S, double L) RgbToHsl(Color c)
    {
        double r = c.R / 255.0, g = c.G / 255.0, b = c.B / 255.0;
        double max = Math.Max(r, Math.Max(g, b));
        double min = Math.Min(r, Math.Min(g, b));
        double l = (max + min) / 2.0;
        double h = 0, s = 0;

        if (max != min)
        {
            double d = max - min;
            s = l > 0.5 ? d / (2.0 - max - min) : d / (max + min);

            if (max == r)
                h = ((g - b) / d + (g < b ? 6 : 0)) * 60.0;
            else if (max == g)
                h = ((b - r) / d + 2) * 60.0;
            else
                h = ((r - g) / d + 4) * 60.0;
        }

        return (h, s, l);
    }

    private static Color HslToRgb(double h, double s, double l)
    {
        if (s == 0)
        {
            int v = (int)Math.Round(l * 255);
            return Color.FromArgb(v, v, v);
        }

        double q = l < 0.5 ? l * (1 + s) : l + s - l * s;
        double p = 2 * l - q;
        double hNorm = h / 360.0;

        double r = HueToRgb(p, q, hNorm + 1.0 / 3.0);
        double g = HueToRgb(p, q, hNorm);
        double b = HueToRgb(p, q, hNorm - 1.0 / 3.0);

        return Color.FromArgb(
            (int)Math.Round(r * 255),
            (int)Math.Round(g * 255),
            (int)Math.Round(b * 255));
    }

    private static double HueToRgb(double p, double q, double t)
    {
        if (t < 0) t += 1;
        if (t > 1) t -= 1;
        if (t < 1.0 / 6.0) return p + (q - p) * 6 * t;
        if (t < 1.0 / 2.0) return q;
        if (t < 2.0 / 3.0) return p + (q - p) * (2.0 / 3.0 - t) * 6;
        return p;
    }

    /// <summary>
    /// Blend a palette color at ~30% alpha with a dark background.
    /// </summary>
    public static Color BlendWithBackground(Color paletteColor, Color background)
    {
        const float alpha = 0.50f;
        return Color.FromArgb(
            (int)(background.R + (paletteColor.R - background.R) * alpha),
            (int)(background.G + (paletteColor.G - background.G) * alpha),
            (int)(background.B + (paletteColor.B - background.B) * alpha));
    }

    /// <summary>
    /// Synonym dictionary mapping adjective/alternate forms to canonical nouns.
    /// </summary>
    private static readonly Dictionary<string, string> Synonyms = new(StringComparer.OrdinalIgnoreCase)
    {
        // Anatomical
        { "hepatic", "liver" }, { "hepatomegaly", "liver" },
        { "renal", "kidney" }, { "nephro", "kidney" },
        { "pulmonary", "lung" }, { "pulmonic", "lung" },
        { "cardiac", "heart" }, { "myocardial", "heart" }, { "cardiomegaly", "heart" },
        { "cerebral", "brain" }, { "intracranial", "brain" }, { "cranial", "brain" },
        { "osseous", "bone" }, { "skeletal", "bone" }, { "bony", "bone" },
        { "splenic", "spleen" }, { "splenomegaly", "spleen" },
        { "pancreatic", "pancreas" },
        { "aortic", "aorta" },
        { "pleural", "pleura" },
        { "colonic", "colon" },
        { "biliary", "bile duct" }, { "choledochal", "bile duct" },
        { "thyroid", "thyroid" }, { "thyroidal", "thyroid" },
        { "adrenal", "adrenal gland" }, { "suprarenal", "adrenal gland" },
        { "prostatic", "prostate" },
        { "uterine", "uterus" }, { "endometrial", "uterus" },
        { "ovarian", "ovary" },
        { "vesical", "bladder" },
        { "gastric", "stomach" },
        { "duodenal", "duodenum" },
        { "ileal", "ileum" }, { "jejunal", "jejunum" },
        { "cecal", "cecum" }, { "appendiceal", "appendix" },
        { "rectal", "rectum" },
        { "esophageal", "esophagus" },
        { "tracheal", "trachea" },
        { "bronchial", "bronchus" },
        { "vertebral", "vertebra" }, { "spinal", "spine" },
        { "pericardial", "pericardium" },
        { "peritoneal", "peritoneum" },
        { "retroperitoneal", "retroperitoneum" },
        { "mediastinal", "mediastinum" },
        { "hilar", "hilum" },
        { "mesenteric", "mesentery" },
        { "inguinal", "groin" },
        { "femoral", "femur" },
        { "humeral", "humerus" },
        { "tibial", "tibia" },
        { "pelvic", "pelvis" },

        // Pathology
        { "nephrolithiasis", "stone" }, { "urolithiasis", "stone" },
        { "cholelithiasis", "gallstone" }, { "gallbladder stone", "gallstone" },
        { "calculus", "stone" }, { "calculi", "stone" }, { "calcified", "calcification" },
        { "thrombus", "clot" }, { "thrombosis", "clot" }, { "thrombotic", "clot" },
        { "embolus", "clot" }, { "embolism", "clot" }, { "embolic", "clot" }, { "filling defect", "clot" },
        { "hemorrhage", "bleeding" }, { "hemorrhagic", "bleeding" },
        { "hematoma", "bleeding" },
        { "edema", "swelling" }, { "edematous", "swelling" },
        { "stenosis", "narrowing" }, { "stenotic", "narrowing" },
        { "hernia", "herniation" }, { "herniated", "herniation" },
        { "fracture", "fracture" }, { "fractured", "fracture" },
        { "effusion", "effusion" },
        { "consolidation", "consolidation" }, { "consolidated", "consolidation" },
        { "atelectasis", "atelectasis" }, { "atelectatic", "atelectasis" },
        { "pneumothorax", "pneumothorax" },
        { "pneumonia", "pneumonia" },
        { "abscess", "abscess" },
        { "mass", "mass" }, { "lesion", "lesion" }, { "nodule", "nodule" }, { "nodular", "nodule" },
        { "tumor", "tumor" }, { "neoplasm", "tumor" }, { "neoplastic", "tumor" },
        { "malignant", "malignancy" }, { "malignancy", "malignancy" },
        { "metastasis", "metastatic disease" }, { "metastatic", "metastatic disease" }, { "metastases", "metastatic disease" },
        { "lymphadenopathy", "lymph node enlargement" }, { "lymph node", "lymph node" },
        { "aneurysm", "aneurysm" }, { "aneurysmal", "aneurysm" },
        { "dissection", "dissection" },
        { "obstruction", "obstruction" }, { "obstructing", "obstruction" }, { "obstructive", "obstruction" },
        { "dilatation", "dilation" }, { "dilated", "dilation" }, { "dilation", "dilation" },
        { "hydronephrosis", "hydronephrosis" }, { "hydroureter", "hydroureter" },
        { "ascites", "ascites" },
        { "swollen", "swelling" },

        // Radiology descriptive
        { "opacification", "opacity" }, { "opacified", "opacity" }, { "opaque", "opacity" },
        { "lucent", "lucency" }, { "radiolucent", "lucency" },
        { "sclerotic", "sclerosis" },
        { "osteolytic", "lytic" },

        // Degenerative / MSK
        { "degenerative", "degeneration" }, { "spondylosis", "degeneration" },
        { "arthritic", "arthritis" }, { "arthropathy", "arthritis" },
        { "scoliotic", "scoliosis" },
        { "kyphotic", "kyphosis" },
        { "lordotic", "lordosis" },

        // Vascular
        { "ectatic", "dilation" }, { "ectasia", "dilation" },
        { "tortuous", "tortuosity" },
        { "atherosclerotic", "atherosclerosis" }, { "atheromatous", "atherosclerosis" },

        // Tissue changes
        { "fibrotic", "fibrosis" },
        { "emphysematous", "emphysema" },
        { "necrotic", "necrosis" },
        { "ischemic", "ischemia" },
        { "infarcted", "infarction" }, { "infarct", "infarction" },

        // Structural
        { "perforated", "perforation" },
        { "displaced", "displacement" },
        { "compressed", "compression" },
        { "distended", "distension" },
        { "collapsed", "collapse" },

        // Neoplastic
        { "carcinoma", "malignancy" }, { "carcinomatous", "malignancy" },

        // Size/shape descriptors
        { "thickened", "thickening" },
        { "enlarged", "enlargement" },
        { "prominent", "prominence" },

        // Short but important pathology terms (< 5 chars, need synonym to be included)
        { "cystic", "cyst" }, { "cyst", "cyst" },

        // Missing anatomical adjective → noun
        { "abdominal", "abdomen" },
        { "sacral", "sacrum" },
    };

    /// <summary>
    /// Laterality terms that alone are too generic to establish a correlation match.
    /// These still contribute to scoring (to pick left vs right impression) but
    /// at least one non-laterality term must also match.
    /// </summary>
    private static readonly HashSet<string> LateralityTerms = new(StringComparer.OrdinalIgnoreCase)
    {
        "left", "right"
    };

    /// <summary>
    /// Anatomical terms for direct matching (these are already canonical).
    /// </summary>
    private static readonly HashSet<string> AnatomicalTerms = new(StringComparer.OrdinalIgnoreCase)
    {
        "liver", "kidney", "lung", "brain", "heart", "spleen", "pancreas", "aorta",
        "pleura", "colon", "bile duct", "thyroid", "adrenal", "prostate", "uterus",
        "ovary", "bladder", "stomach", "duodenum", "ileum", "jejunum", "cecum",
        "appendix", "rectum", "esophagus", "trachea", "bronchus", "vertebra", "spine",
        "pericardium", "peritoneum", "retroperitoneum", "mediastinum", "hilum",
        "mesentery", "femur", "humerus", "tibia", "pelvis", "gallbladder",
        "bowel", "small bowel", "large bowel", "chest wall", "diaphragm",
        "bone", "lymph node",
        // Laterality and short anatomical terms (< 5 chars, need explicit listing)
        "left", "right", "lobe", "rib", "disc", "disk",
        "atrium", "ventricle", "artery", "vein", "sacrum"
    };

    /// <summary>
    /// Correlate findings and impression sections of a radiology report.
    /// </summary>
    public static CorrelationResult Correlate(string reportText, int? seed = null)
    {
        var result = new CorrelationResult { Source = "heuristic" };

        Logger.Trace($"Correlate: input length={reportText.Length}");

        var contextTerms = ExtractExamContextTerms(reportText);
        var (findingsText, impressionText) = ExtractSections(reportText);
        Logger.Trace($"Correlate: findings={findingsText.Length} chars, impression={impressionText.Length} chars");
        if (string.IsNullOrWhiteSpace(findingsText) || string.IsNullOrWhiteSpace(impressionText))
        {
            Logger.Trace("Correlate: Missing FINDINGS or IMPRESSION section");
            return result;
        }

        var impressionItems = ParseImpressionItems(impressionText);
        var findingsSentences = ParseFindingsSentences(findingsText);

        Logger.Trace($"Correlate: {impressionItems.Count} impression items, {findingsSentences.Count} findings sentences");

        if (impressionItems.Count == 0 || findingsSentences.Count == 0)
            return result;

        // Group findings sentences by subsection (contiguous blocks)
        var findingsBlocks = GroupFindingsBySubsection(findingsText);
        Logger.Trace($"Correlate: {findingsBlocks.Count} findings subsection blocks");

        int colorIndex = 0;
        foreach (var (index, text) in impressionItems)
        {
            // Skip negative/normal impression statements — these match too broadly
            if (NegativeSentencePattern.IsMatch(text))
            {
                Logger.Trace($"Correlate: Impression #{index} skipped (negative statement): {text.Substring(0, Math.Min(80, text.Length))}");
                continue;
            }

            var impressionTerms = ExtractTerms(text, contextTerms);
            Logger.Trace($"Correlate: Impression #{index} terms: [{string.Join(", ", impressionTerms)}] from: {text.Substring(0, Math.Min(80, text.Length))}");

            if (impressionTerms.Count == 0) continue;

            // Find the best matching subsection block (most shared terms)
            // This prevents a single common word from spreading across the whole report
            int bestScore = 0;
            int bestSubstantiveScore = 0;
            (string? Header, List<string> Sentences)? bestBlock = null;

            foreach (var block in findingsBlocks)
            {
                // Pool all terms from the block's sentences
                var blockTerms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var sentence in block.Sentences)
                {
                    foreach (var t in ExtractTerms(sentence, contextTerms))
                        blockTerms.Add(t);
                }

                var shared = impressionTerms.Intersect(blockTerms, StringComparer.OrdinalIgnoreCase).ToList();
                int substantiveCount = shared.Count(t => !LateralityTerms.Contains(t));
                if (shared.Count > bestScore)
                {
                    bestScore = shared.Count;
                    bestSubstantiveScore = substantiveCount;
                    bestBlock = block;
                    Logger.Trace($"Correlate:   Block match score={shared.Count} substantive={substantiveCount} terms=[{string.Join(", ", shared)}]");
                }
            }

            // Require at least 1 non-laterality shared term with the best block
            if (bestSubstantiveScore >= 1 && bestBlock != null)
            {
                // Only include sentences from the best block that actually share substantive terms
                var matchedSentences = new List<string>();
                foreach (var sentence in bestBlock.Value.Sentences)
                {
                    var sentenceTerms = ExtractTerms(sentence, contextTerms);
                    var sentenceShared = impressionTerms.Intersect(sentenceTerms, StringComparer.OrdinalIgnoreCase).ToList();
                    if (sentenceShared.Any(t => !LateralityTerms.Contains(t)))
                    {
                        matchedSentences.Add(sentence);
                    }
                }

                if (matchedSentences.Count > 0)
                {
                    var category = GetCategoryForHeader(bestBlock.Value.Header);
                    category = RefineCategoryFromContent(category, matchedSentences);
                    result.Items.Add(new CorrelationItem
                    {
                        ImpressionIndex = index,
                        ImpressionText = text,
                        MatchedFindings = matchedSentences,
                        ColorIndex = colorIndex % Palette.Length,
                        HighlightColor = GetColorForCategory(category)
                    });
                    colorIndex++;
                }
            }
        }

        DiversifyCategoryColors(result.Items);
        return result;
    }

    /// <summary>
    /// Regex pattern matching negative/normal/incidental finding sentences that don't need impression representation.
    /// Start-anchored patterns prevent filtering mixed sentences like
    /// "The kidney shows no hydronephrosis but a new 2 cm mass."
    /// Unanchored patterns match anywhere for unambiguous indicators.
    /// </summary>
    private static readonly Regex NegativeSentencePattern = new(
        @"^\s*(?:no\s|normal\b|negative\b|there\s+(?:is|are)\s+no\s|not\s)" +
        @"|\bunremarkable\b|\bwithin\s+normal\b" +
        @"|\b(?:was|has\s+been)\s+performed\b" +               // surgical history
        @"|\bdemonstrates?\s+no\s" +                            // mid-sentence negative
        @"|\b(?:is|are)\s+normal\b" +                            // "is normal", "is normal in size/caliber"
        @"|\bnot\s+clearly\s+(?:seen|visualized|identified)\b" + // non-visualization
        @"|\b(?:is|are)\s+clear\b" +                            // "airways are clear"
        @"|\bno\s+evidence\b",                                  // "no evidence of acute..."
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Reversed correlation: starts from dictated findings and checks whether each
    /// is represented in the impression. Unmatched dictated findings get a unique
    /// "orphan" color that appears nowhere in the impression.
    /// </summary>
    public static CorrelationResult CorrelateReversed(string reportText, HashSet<string>? dictatedSentences = null, int? seed = null)
    {
        var result = new CorrelationResult { Source = "heuristic-reversed" };

        Logger.Trace($"CorrelateReversed: input length={reportText.Length}, dictatedSentences={dictatedSentences?.Count ?? -1}");

        var contextTerms = ExtractExamContextTerms(reportText);
        var (findingsText, impressionText) = ExtractSections(reportText);
        if (string.IsNullOrWhiteSpace(findingsText) || string.IsNullOrWhiteSpace(impressionText))
        {
            Logger.Trace("CorrelateReversed: Missing FINDINGS or IMPRESSION section");
            return result;
        }

        var findingsSentences = ParseFindingsSentences(findingsText);
        var impressionItems = ParseImpressionItems(impressionText);

        if (findingsSentences.Count == 0 || impressionItems.Count == 0)
            return result;

        // Filter findings to only dictated sentences (if baseline was available)
        List<string> filteredFindings;
        if (dictatedSentences != null && dictatedSentences.Count > 0)
        {
            filteredFindings = new List<string>();
            foreach (var sentence in findingsSentences)
            {
                var normalized = Regex.Replace(sentence.ToLowerInvariant().Trim(), @"\s+", " ");
                if (dictatedSentences.Contains(normalized))
                    filteredFindings.Add(sentence);
            }
            Logger.Trace($"CorrelateReversed: {filteredFindings.Count}/{findingsSentences.Count} findings are dictated");
        }
        else
        {
            filteredFindings = new List<string>(findingsSentences);
        }

        // Identify negative/normal statements (these can match impressions but won't become orphans)
        // When dictatedSentences is provided, we KNOW which findings are dictated — negative
        // findings that were confirmed dictated should still be highlighted (user said them),
        // they just shouldn't become orphan warnings.
        var negativeFindings = new HashSet<string>(StringComparer.Ordinal);
        bool hasDictatedFilter = dictatedSentences != null && dictatedSentences.Count > 0;
        foreach (var sentence in filteredFindings)
        {
            if (NegativeSentencePattern.IsMatch(sentence))
            {
                negativeFindings.Add(sentence);
                Logger.Trace($"CorrelateReversed: Negative finding{(hasDictatedFilter ? " (dictated, will highlight+orphan)" : " (template, won't orphan)")}: {sentence.Substring(0, Math.Min(60, sentence.Length))}");
            }
        }
        int significantCount = filteredFindings.Count - negativeFindings.Count;
        Logger.Trace($"CorrelateReversed: {significantCount} significant + {negativeFindings.Count} negative findings (dictatedFilter={hasDictatedFilter})");

        // Safety net: if no dictated sentences provided AND most findings are significant (suggesting
        // baseline wasn't captured), fall back to the old impression-first approach
        if (dictatedSentences == null && significantCount > findingsSentences.Count * 0.8)
        {
            Logger.Trace("CorrelateReversed: No baseline, most findings pass filter — falling back to Correlate()");
            return Correlate(reportText, seed);
        }

        // When we have a dictated filter, negative findings count as significant
        // (they're confirmed dictated content that should be correlated)
        if (significantCount == 0 && !hasDictatedFilter)
            return result;

        // Build subsection blocks so we can look up header for each finding sentence
        var findingsBlocks = GroupFindingsBySubsection(findingsText);

        // Build a lookup: finding sentence → subsection header
        var sentenceToHeader = new Dictionary<string, string?>(StringComparer.Ordinal);
        foreach (var (header, sentences) in findingsBlocks)
        {
            foreach (var sentence in sentences)
            {
                sentenceToHeader.TryAdd(sentence, header);
            }
        }

        // For each finding, find the best-matching impression item
        // All findings (including negative) participate in matching,
        // but only positive findings become orphans when unmatched
        var matchedGroups = new Dictionary<int, (string impressionText, List<string> findings)>();
        var orphanFindings = new List<string>();

        foreach (var finding in filteredFindings)
        {
            bool isNegative = negativeFindings.Contains(finding);
            var findingTerms = ExtractTerms(finding, contextTerms);
            Logger.Trace($"CorrelateReversed: Finding terms=[{string.Join(", ", findingTerms)}] neg={isNegative} for: {finding.Substring(0, Math.Min(60, finding.Length))}");
            if (findingTerms.Count == 0)
            {
                if (!isNegative) orphanFindings.Add(finding);
                continue;
            }

            int bestScore = 0;
            int bestSubstantiveScore = 0;
            int bestImpressionIdx = -1;
            string bestImpressionText = "";

            foreach (var (index, text) in impressionItems)
            {
                // Skip negative/normal impression statements — these match too broadly
                if (NegativeSentencePattern.IsMatch(text))
                    continue;

                var impressionTerms = ExtractTerms(text, contextTerms);
                var shared = findingTerms.Intersect(impressionTerms, StringComparer.OrdinalIgnoreCase).ToList();
                int substantiveCount = shared.Count(t => !LateralityTerms.Contains(t));
                Logger.Trace($"CorrelateReversed:   vs Impression #{index} terms=[{string.Join(", ", impressionTerms)}] shared=[{string.Join(", ", shared)}] score={shared.Count} substantive={substantiveCount}");
                if (shared.Count > bestScore)
                {
                    bestScore = shared.Count;
                    bestSubstantiveScore = substantiveCount;
                    bestImpressionIdx = index;
                    bestImpressionText = text;
                }
            }

            // Require at least one non-laterality shared term to count as a match
            if (bestSubstantiveScore >= 1 && bestImpressionIdx >= 0)
            {
                Logger.Trace($"CorrelateReversed: MATCHED finding to Impression #{bestImpressionIdx} (score={bestScore}, substantive={bestSubstantiveScore}, neg={isNegative})");
                // When we have a dictated filter, negative findings are confirmed dictated content
                // and should be highlighted. Without a filter, exclude negatives to avoid
                // highlighting boilerplate template content.
                if (!isNegative || hasDictatedFilter)
                {
                    if (!matchedGroups.ContainsKey(bestImpressionIdx))
                        matchedGroups[bestImpressionIdx] = (bestImpressionText, new List<string>());
                    matchedGroups[bestImpressionIdx].findings.Add(finding);
                }
            }
            else if (!isNegative || hasDictatedFilter)
            {
                // Orphan: positive finding, or dictated pertinent negative with no impression match
                Logger.Trace($"CorrelateReversed: ORPHAN finding (bestScore={bestScore}, substantive={bestSubstantiveScore}, neg={isNegative}, dictated={hasDictatedFilter})");
                orphanFindings.Add(finding);
            }
            else
            {
                Logger.Trace($"CorrelateReversed: Negative template finding unmatched, not orphaned");
            }
        }

        // Build result: matched groups get a color based on body-part category
        int colorIndex = 0;
        foreach (var (groupIdx, (groupText, groupFindings)) in matchedGroups)
        {
            // Determine color from the first finding's subsection header
            string? header = null;
            foreach (var f in groupFindings)
            {
                if (sentenceToHeader.TryGetValue(f, out header) && header != null)
                    break;
            }
            var category = GetCategoryForHeader(header);
            category = RefineCategoryFromContent(category, groupFindings);

            result.Items.Add(new CorrelationItem
            {
                ImpressionIndex = groupIdx,
                ImpressionText = groupText,
                MatchedFindings = groupFindings,
                ColorIndex = colorIndex % Palette.Length,
                HighlightColor = GetColorForCategory(category)
            });
            colorIndex++;
        }

        // Orphan findings: color from body-part category
        foreach (var orphan in orphanFindings)
        {
            sentenceToHeader.TryGetValue(orphan, out var orphanHeader);
            var category = GetCategoryForHeader(orphanHeader);
            category = RefineCategoryFromContent(category, new List<string> { orphan });

            result.Items.Add(new CorrelationItem
            {
                ImpressionIndex = -1,
                ImpressionText = "",
                MatchedFindings = new List<string> { orphan },
                ColorIndex = colorIndex % Palette.Length,
                HighlightColor = GetColorForCategory(category)
            });
            colorIndex++;
        }

        DiversifyCategoryColors(result.Items);
        Logger.Trace($"CorrelateReversed: {matchedGroups.Count} matched groups, {orphanFindings.Count} orphans");
        return result;
    }

    /// <summary>
    /// Extract context terms from the EXAM description line to exclude from correlation.
    /// Terms like "lumbar", "spine" in a lumbar spine study appear everywhere and don't indicate pathology.
    /// </summary>
    internal static HashSet<string> ExtractExamContextTerms(string reportText)
    {
        var contextTerms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Find EXAM: line (inline "EXAM: description" or "EXAM:" followed by description on next line)
        var lines = reportText.Replace("\r\n", "\n").Split('\n');
        string? examLine = null;
        bool sawExamHeader = false;

        foreach (var line in lines)
        {
            var trimmed = line.Trim();
            if (trimmed.StartsWith("EXAM:", StringComparison.OrdinalIgnoreCase))
            {
                var rest = trimmed.Substring(5).Trim();
                if (!string.IsNullOrEmpty(rest))
                {
                    examLine = rest;
                    break;
                }
                // Standalone "EXAM:" — description is on the next line
                sawExamHeader = true;
                continue;
            }
            if (sawExamHeader && !string.IsNullOrWhiteSpace(trimmed))
            {
                // Don't grab another section header as the exam description
                if (Regex.IsMatch(trimmed, @"^[A-Z\s]+:\s*$"))
                    break;
                examLine = trimmed;
                break;
            }
        }

        if (string.IsNullOrWhiteSpace(examLine)) return contextTerms;

        // Strip date/time suffix (e.g., "02/06/2026 06:59:00 PM")
        examLine = Regex.Replace(examLine, @"\d{1,2}/\d{1,2}/\d{4}.*$", "").Trim();

        // Extract significant words from exam description
        var words = Regex.Split(examLine.ToLowerInvariant(), @"[^a-z]+")
            .Where(w => w.Length >= 4) // Only meaningful words
            .ToArray();

        foreach (var word in words)
        {
            // Skip very common modifiers that aren't anatomical context
            if (word is "with" or "without" or "views" or "view" or "contrast" or "portable")
                continue;
            contextTerms.Add(word);
            // Also add synonym-mapped forms
            if (Synonyms.TryGetValue(word, out var canonical))
                contextTerms.Add(canonical);
            if (AnatomicalTerms.Contains(word))
                contextTerms.Add(word);
        }

        if (contextTerms.Count > 0)
            Logger.Trace($"Exam context terms (excluded from matching): [{string.Join(", ", contextTerms)}]");

        return contextTerms;
    }

    /// <summary>
    /// Extract FINDINGS and IMPRESSION sections from report text.
    /// </summary>
    internal static (string findings, string impression) ExtractSections(string text)
    {
        // Normalize
        text = text.Replace("\r\n", "\n").Replace("\r", "\n");

        // Find section boundaries using regex
        var findingsMatch = Regex.Match(text, @"^FINDINGS:\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        var impressionMatch = Regex.Match(text, @"^IMPRESSION:\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);

        if (!findingsMatch.Success || !impressionMatch.Success)
        {
            // Try inline format: "FINDINGS: content" on same line
            findingsMatch = Regex.Match(text, @"FINDINGS:", RegexOptions.IgnoreCase);
            impressionMatch = Regex.Match(text, @"IMPRESSION:", RegexOptions.IgnoreCase);
        }

        if (!findingsMatch.Success || !impressionMatch.Success)
            return ("", "");

        // Ensure FINDINGS comes before IMPRESSION
        if (findingsMatch.Index >= impressionMatch.Index)
            return ("", "");

        string findings = text.Substring(
            findingsMatch.Index + findingsMatch.Length,
            impressionMatch.Index - findingsMatch.Index - findingsMatch.Length).Trim();

        string impression = text.Substring(
            impressionMatch.Index + impressionMatch.Length).Trim();

        // Trim impression at next major section if any
        var nextSection = Regex.Match(impression, @"^\s*[A-Z][A-Z\s]+:", RegexOptions.Multiline);
        if (nextSection.Success && nextSection.Index > 0)
        {
            impression = impression.Substring(0, nextSection.Index).Trim();
        }

        return (findings, impression);
    }

    /// <summary>
    /// Parse impression text into numbered items. If no numbered items, treat whole text as one item.
    /// </summary>
    private static List<(int index, string text)> ParseImpressionItems(string impressionText)
    {
        var items = new List<(int, string)>();

        // Match numbered items: "1. text" or "1) text"
        var matches = Regex.Matches(impressionText, @"(?:^|\n)\s*(\d+)[.)]\s*(.+?)(?=\n\s*\d+[.)]|\n\s*$|$)", RegexOptions.Singleline);

        if (matches.Count > 0)
        {
            foreach (Match m in matches)
            {
                int num = int.Parse(m.Groups[1].Value);
                string text = m.Groups[2].Value.Trim();
                if (!string.IsNullOrWhiteSpace(text))
                    items.Add((num, text));
            }
        }
        else
        {
            // No numbered items - treat entire impression as one item
            if (!string.IsNullOrWhiteSpace(impressionText))
                items.Add((1, impressionText.Trim()));
        }

        return items;
    }

    /// <summary>
    /// Parse findings into individual sentences, splitting on period + space or newlines.
    /// Strip subsection headers to get content only.
    /// </summary>
    private static List<string> ParseFindingsSentences(string findingsText)
    {
        var sentences = new List<string>();

        // Split on newlines first to separate subsections
        var lines = findingsText.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line)) continue;

            // Strip subsection header prefix (e.g., "LUNGS AND PLEURA: text" -> "text")
            string content = StripSubsectionHeader(line);
            if (string.IsNullOrWhiteSpace(content)) continue;

            // Don't include pure section headers
            if (IsAllCapsHeader(content)) continue;

            // Split content into sentences on ". " boundary
            var sentenceParts = Regex.Split(content, @"(?<=\.)\s+");
            foreach (var part in sentenceParts)
            {
                var trimmed = part.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed) && trimmed.Length >= 3)
                    sentences.Add(trimmed);
            }
        }

        return sentences;
    }

    /// <summary>
    /// Group findings text into subsection blocks. Each block is a list of sentences
    /// under the same subsection header. Sentences without a subsection header form their own block.
    /// This ensures correlation matches contiguous related sentences, not scattered ones.
    /// Returns the header text (before colon) for each block, or null if no header was found.
    /// </summary>
    private static List<(string? Header, List<string> Sentences)> GroupFindingsBySubsection(string findingsText)
    {
        var blocks = new List<(string? Header, List<string> Sentences)>();
        var currentBlock = new List<string>();
        string? currentHeader = null;

        var lines = findingsText.Split('\n', StringSplitOptions.RemoveEmptyEntries);

        foreach (var rawLine in lines)
        {
            var line = rawLine.Trim();
            if (string.IsNullOrWhiteSpace(line)) continue;

            // Check if this line starts a new subsection (ALL CAPS header with colon)
            bool isHeader = false;
            string? detectedHeader = null;
            int colonIdx = line.IndexOf(':');
            if (colonIdx > 0 && colonIdx < line.Length - 1)
            {
                var prefix = line.Substring(0, colonIdx);
                if (IsAllCapsHeader(prefix))
                {
                    isHeader = true;
                    detectedHeader = prefix.Trim();
                }
            }
            else if (IsAllCapsHeader(line))
            {
                isHeader = true;
                detectedHeader = line.TrimEnd(':').Trim();
            }

            if (isHeader && currentBlock.Count > 0)
            {
                // Save previous block, start new one
                blocks.Add((currentHeader, currentBlock));
                currentBlock = new List<string>();
            }

            if (isHeader)
                currentHeader = detectedHeader;

            // Strip header prefix and split into sentences
            string content = StripSubsectionHeader(line);
            if (string.IsNullOrWhiteSpace(content) || IsAllCapsHeader(content)) continue;

            var sentenceParts = Regex.Split(content, @"(?<=\.)\s+");
            foreach (var part in sentenceParts)
            {
                var trimmed = part.Trim();
                if (!string.IsNullOrWhiteSpace(trimmed) && trimmed.Length >= 3)
                    currentBlock.Add(trimmed);
            }
        }

        if (currentBlock.Count > 0)
            blocks.Add((currentHeader, currentBlock));

        // If no subsection headers were found (simple reports), treat all sentences as one block
        if (blocks.Count == 0 && findingsText.Length > 0)
        {
            var allSentences = ParseFindingsSentences(findingsText);
            if (allSentences.Count > 0)
                blocks.Add((null, allSentences));
        }

        return blocks;
    }

    private static string StripSubsectionHeader(string line)
    {
        int colonIdx = line.IndexOf(':');
        if (colonIdx <= 0 || colonIdx >= line.Length - 1) return line;

        var prefix = line.Substring(0, colonIdx);
        if (IsAllCapsHeader(prefix))
            return line.Substring(colonIdx + 1).Trim();

        return line;
    }

    private static bool IsAllCapsHeader(string text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        var trimmed = text.Trim().TrimEnd(':');
        if (trimmed.Length < 2) return false;

        int upper = 0, letters = 0;
        foreach (char c in trimmed)
        {
            if (char.IsLetter(c))
            {
                letters++;
                if (char.IsUpper(c)) upper++;
            }
        }
        return letters > 0 && (double)upper / letters > 0.8;
    }

    /// <summary>
    /// Common English stopwords to exclude from content word matching.
    /// </summary>
    private static readonly HashSet<string> StopWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "about", "above", "after", "again", "along", "also", "appears", "are",
        "been", "before", "being", "below", "between", "both", "cannot",
        "change", "changes", "compatible", "compared", "could",
        "demonstrate", "demonstrates", "demonstrated", "does",
        "each", "either", "evaluate", "evaluation", "evidence",
        "findings", "following", "from", "given", "have", "having",
        "identified", "impression", "including", "interval",
        "large", "likely", "limited", "measure", "measures", "measuring",
        "mild", "mildly", "minimal", "moderate", "moderately",
        "neither", "normal", "noted", "number",
        "other", "otherwise", "overall",
        "partially", "patient", "please", "possible", "possibly",
        "prior", "probably",
        "redemonstrated", "related", "remain", "remains",
        "seen", "series", "severe", "severely", "several", "should",
        "since", "small", "stable", "status", "study", "suggest", "suggests",
        "there", "these", "those", "through", "total",
        "unchanged", "under", "unremarkable",
        "visualized", "well", "where", "which", "within", "without"
    };

    /// <summary>
    /// Simple depluralization: strips trailing 's' or 'es' to normalize plurals.
    /// E.g., "tissues" → "tissue", "masses" → "mass", "calculi" handled by synonyms.
    /// </summary>
    private static string Depluralize(string word)
    {
        if (word.Length <= 3) return word;
        if (word.EndsWith("ses") || word.EndsWith("zes") || word.EndsWith("xes") || word.EndsWith("ches") || word.EndsWith("shes"))
            return word.Substring(0, word.Length - 2);
        if (word.EndsWith("ies"))
            return word.Substring(0, word.Length - 3) + "y";
        if (word.EndsWith("s") && !word.EndsWith("ss") && !word.EndsWith("us") && !word.EndsWith("is"))
            return word.Substring(0, word.Length - 1);
        return word;
    }

    /// <summary>
    /// Extract canonical medical terms from a text fragment.
    /// Returns a set of normalized terms for matching.
    /// Includes: synonym-mapped terms, anatomical terms, and significant content words (5+ chars).
    /// </summary>
    internal static HashSet<string> ExtractTerms(string text) => ExtractTerms(text, null);

    internal static HashSet<string> ExtractTerms(string text, HashSet<string>? contextExclusions)
    {
        var terms = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(text)) return terms;

        // Tokenize: split on non-letter characters
        var words = Regex.Split(text.ToLowerInvariant(), @"[^a-z]+")
            .Where(w => w.Length >= 3)
            .ToArray();

        foreach (var word in words)
        {
            // Check synonym dictionary (try both raw word and depluralized form)
            var singular = Depluralize(word);
            if (Synonyms.TryGetValue(word, out var canonical))
            {
                terms.Add(canonical);
            }
            else if (singular != word && Synonyms.TryGetValue(singular, out var canonical2))
            {
                terms.Add(canonical2);
            }

            // Check if word itself is an anatomical term
            if (AnatomicalTerms.Contains(word))
            {
                terms.Add(word);
            }
            else if (singular != word && AnatomicalTerms.Contains(singular))
            {
                terms.Add(singular);
            }

            // Include significant content words (5+ chars, not stopwords)
            // This catches medical terms not in our dictionaries
            // Use depluralized form so "tissues" and "tissue" match
            if (word.Length >= 5 && !StopWords.Contains(word))
            {
                terms.Add(singular);
            }
        }

        // Also check multi-word anatomical terms
        var lowerText = text.ToLowerInvariant();
        foreach (var term in AnatomicalTerms)
        {
            if (term.Contains(' ') && lowerText.Contains(term))
            {
                terms.Add(term);
            }
        }

        // Check multi-word synonyms
        foreach (var kvp in Synonyms)
        {
            if (kvp.Key.Contains(' ') && lowerText.Contains(kvp.Key.ToLowerInvariant()))
            {
                terms.Add(kvp.Value);
            }
        }

        // Remove exam context terms (e.g., "lumbar", "spine" in a lumbar spine study)
        if (contextExclusions != null && contextExclusions.Count > 0)
        {
            terms.ExceptWith(contextExclusions);
        }

        return terms;
    }
}
