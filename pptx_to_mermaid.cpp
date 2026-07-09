// pptx_to_mermaid.cpp
// Convert box diagrams in PowerPoint (.pptx) files to Mermaid flowchart syntax.
//
// Build (MinGW / Git Bash on Windows):
//   g++ -std=c++17 -O2 -Ivendor pptx_to_mermaid.cpp vendor/tinyxml2.cpp -o pptx_to_mermaid.exe -lz
//
// Build (MSVC, Developer Command Prompt):
//   cl /std:c++17 /EHsc /Ivendor pptx_to_mermaid.cpp vendor\tinyxml2.cpp
//
// Build (Linux):
//   g++ -std=c++17 -O2 -Ivendor pptx_to_mermaid.cpp vendor/tinyxml2.cpp -o pptx_to_mermaid -lz
//
// Usage:
//   pptx_to_mermaid input.pptx [output.md] [--slides 1 2 3] [--direction TD|LR|auto]

#include "tinyxml2.h"
#include "miniz.h"

#include <algorithm>
#include <cctype>
#include <cstdint>
#include <cstdio>
#include <cstring>
#include <filesystem>
#include <fstream>
#include <iostream>
#include <map>
#include <regex>
#include <set>
#include <sstream>
#include <string>
#include <utility>
#include <vector>

namespace fs = std::filesystem;

// ---------------------------------------------------------------------------
// Namespaces used by OOXML / DrawingML.
// ---------------------------------------------------------------------------
namespace oox {
constexpr const char* A = "http://schemas.openxmlformats.org/drawingml/2006/main";
constexpr const char* P = "http://schemas.openxmlformats.org/presentationml/2006/main";
constexpr const char* R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
constexpr const char* DGM = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
} // namespace oox

// ---------------------------------------------------------------------------
// Data model.
// ---------------------------------------------------------------------------
struct Shape {
    std::string id;
    std::string name;
    std::string text;
    // Per-paragraph text and indent level (from <a:pPr lvl="N"/>). The
    // concatenated `text` is also kept around for the Mermaid path so we
    // don't need to flatten there.
    std::vector<std::pair<int, std::string>> paragraphs;
    int64_t x = 0;
    int64_t y = 0;
    int64_t w = 0;
    int64_t h = 0;
    bool isTitle = false;
    bool isBody = false;
    int phIdx = -1;       // placeholder idx (e.g. 1 = body)
    int phType = 0;       // raw placeholder type value if needed
};

struct Edge {
    std::string from;
    std::string to;
};

struct TableCell {
    std::string text;
};

struct TableRow {
    std::vector<TableCell> cells;
};

struct Table {
    std::string name;
    std::vector<TableRow> rows;
};

struct Picture {
    std::string name;
    std::string relId;        // rId into the slide's rels file
    std::string descr;
};

struct Slide {
    std::vector<Shape> shapes;
    std::vector<Edge> edges;
    std::vector<Table> tables;
    std::vector<Picture> pictures;
};

// A connector as it appears in the XML, before we resolve dangling endpoints.
// `start_id` / `end_id` may be empty when the corresponding side of the
// connector is anchored to a fixed point in space rather than another shape.
struct RawConnector {
    std::string start_id;
    std::string end_id;
    bool bidirectional = false;  // arrowheads on both ends
    int64_t x = 0, y = 0, w = 0, h = 0;  // bounding box of the connector itself
    bool flipH = false, flipV = false;
};

// ---------------------------------------------------------------------------
// ZIP reader backed by miniz.
// ---------------------------------------------------------------------------
class PptxReader {
public:
    ~PptxReader() { close(); }

    bool open(const std::string& path) {
        close();
        memset(&zip_, 0, sizeof(zip_));
        return mz_zip_reader_init_file(&zip_, path.c_str(), 0);
    }

    void close() {
        if (opened_) {
            mz_zip_reader_end(&zip_);
            opened_ = false;
        }
    }

    std::vector<std::string> listFiles(const std::string& prefix = "") const {
        std::vector<std::string> out;
        for (mz_uint i = 0; i < mz_zip_reader_get_num_files(&zip_); ++i) {
            mz_zip_archive_file_stat st;
            if (!mz_zip_reader_file_stat(&zip_, i, &st)) continue;
            std::string name = st.m_filename;
            if (prefix.empty() || name.rfind(prefix, 0) == 0) {
                out.push_back(name);
            }
        }
        std::sort(out.begin(), out.end());
        return out;
    }

    // Returns empty string if entry not found.
    std::string read(const std::string& entry) {
        int idx = mz_zip_reader_locate_file(&zip_, entry.c_str(), nullptr, 0);
        if (idx < 0) return {};
        mz_zip_archive_file_stat st;
        if (!mz_zip_reader_file_stat(&zip_, idx, &st)) return {};
        std::string buf(st.m_uncomp_size, '\0');
        if (!mz_zip_reader_extract_to_mem(&zip_, idx, buf.data(), buf.size(), 0)) {
            return {};
        }
        opened_ = true; // any successful extract means archive is open
        return buf;
    }

private:
    mutable mz_zip_archive zip_{};
    mutable bool opened_ = false;
};

// ---------------------------------------------------------------------------
// XML helpers.
// ---------------------------------------------------------------------------
namespace xmlutil {

// tinyxml2 does not process XML namespaces - element names retain the prefix
// that appears in the source (e.g. "p:sld", "a:t"). We build prefixed names
// directly instead of {namespace}local Clark notation.
std::string qn(const char* prefix, const char* local) {
    std::string s = prefix; s += ":"; s += local;
    return s;
}

const tinyxml2::XMLElement* firstChild(const tinyxml2::XMLElement* e, const char* prefix, const char* local) {
    return e ? e->FirstChildElement(qn(prefix, local).c_str()) : nullptr;
}

// Collect every <a:t> text descendant of `e`, joined by '\n'.
// Paragraphs (<a:p>) become line breaks; multiple runs within a paragraph
// concatenate without a separator.
std::string collectText(const tinyxml2::XMLElement* e) {
    if (!e) return {};
    std::vector<std::string> paragraphs;
    std::string current;
    std::function<void(const tinyxml2::XMLElement*, bool)> walk =
        [&](const tinyxml2::XMLElement* node, bool inParagraph) {
        for (const tinyxml2::XMLElement* c = node->FirstChildElement(); c; c = c->NextSiblingElement()) {
            const char* raw = c->Name();
            std::string name = raw ? raw : "";
            auto colon = name.find(':');
            std::string local = colon == std::string::npos ? name : name.substr(colon + 1);
            if (local == "p") {
                if (!current.empty() || inParagraph) {
                    paragraphs.push_back(current);
                    current.clear();
                }
                walk(c, true);
            } else if (local == "t") {
                if (const char* t = c->GetText()) current += t;
            } else {
                walk(c, inParagraph);
            }
        }
    };
    walk(e, false);
    if (!current.empty() || !paragraphs.empty()) paragraphs.push_back(current);
    while (!paragraphs.empty() && paragraphs.back().empty()) paragraphs.pop_back();
    std::string joined;
    for (size_t i = 0; i < paragraphs.size(); ++i) {
        if (i) joined += '\n';
        joined += paragraphs[i];
    }
    return joined;
}

// Same as collectText but also returns the per-paragraph indent level taken
// from <a:pPr lvl="N"/>. Used to render Markdown bullets that respect the
// progressive hierarchy of nested lists.
std::vector<std::pair<int, std::string>> collectParagraphs(const tinyxml2::XMLElement* e) {
    std::vector<std::pair<int, std::string>> out;
    if (!e) return out;

    // Recursively concatenate every <a:t> text descendant of `node`.
    // Sibling <a:r> runs join seamlessly - only an explicit <a:br> adds a
    // line break - so "模块" + "1" stays as one unit instead of becoming
    // "模块\n1".
    std::function<std::string(const tinyxml2::XMLElement*)> joinText =
        [&](const tinyxml2::XMLElement* node) -> std::string {
        std::string s;
        for (const tinyxml2::XMLElement* c = node->FirstChildElement(); c; c = c->NextSiblingElement()) {
            const char* raw = c->Name();
            std::string name = raw ? raw : "";
            auto colon = name.find(':');
            std::string local = colon == std::string::npos ? name : name.substr(colon + 1);
            if (local == "t") {
                if (auto* t = c->GetText()) s += t;
            } else if (local == "br") {
                s += '\n';
            } else {
                std::string sub = joinText(c);
                if (!sub.empty()) s += sub;
            }
        }
        return s;
    };

    // Walk paragraphs at any depth (handles both direct children and nested
    // cases where the parser didn't surface them).
    std::function<void(const tinyxml2::XMLElement*)> walk =
        [&](const tinyxml2::XMLElement* node) {
        for (const tinyxml2::XMLElement* c = node->FirstChildElement(); c; c = c->NextSiblingElement()) {
            const char* raw = c->Name();
            std::string name = raw ? raw : "";
            auto colon = name.find(':');
            std::string local = colon == std::string::npos ? name : name.substr(colon + 1);
            if (local == "p") {
                int level = 0;
                const tinyxml2::XMLElement* pPr = c->FirstChildElement(qn("a", "pPr").c_str());
                if (pPr) level = pPr->IntAttribute("lvl", 0);
                std::string text = joinText(c);
                out.emplace_back(level, std::move(text));
            } else {
                walk(c);
            }
        }
    };
    walk(e);
    while (!out.empty() && out.back().second.empty()) out.pop_back();
    return out;
}

// Parse xfrm (off + ext). Returns true on success.
bool getXfrm(const tinyxml2::XMLElement* sp_or_cxn, int64_t& x, int64_t& y, int64_t& w, int64_t& h) {
    const tinyxml2::XMLElement* spPr = sp_or_cxn->FirstChildElement(qn("p", "spPr").c_str());
    if (!spPr) {
        // For some elements (graphic frame) the transform lives on a:xfrm directly.
        spPr = sp_or_cxn;
    }
    const tinyxml2::XMLElement* xfrm = spPr->FirstChildElement(qn("a", "xfrm").c_str());
    if (!xfrm) return false;
    const tinyxml2::XMLElement* off = xfrm->FirstChildElement(qn("a", "off").c_str());
    const tinyxml2::XMLElement* ext = xfrm->FirstChildElement(qn("a", "ext").c_str());
    if (!off || !ext) return false;
    x = off->Int64Attribute("x", 0);
    y = off->Int64Attribute("y", 0);
    w = ext->Int64Attribute("cx", 0);
    h = ext->Int64Attribute("cy", 0);
    return true;
}

} // namespace xmlutil

// ---------------------------------------------------------------------------
// Slide parser: walks a single ppt/slides/slideN.xml.
// ---------------------------------------------------------------------------
class SlideParser {
public:
    // Parses one slide and returns everything we care about: shapes,
    // resolved edges, tables, and pictures.
    Slide parse(const std::string& xml) {
        tinyxml2::XMLDocument doc;
        if (doc.Parse(xml.c_str(), xml.size()) != tinyxml2::XML_SUCCESS) return {};
        const tinyxml2::XMLElement* spTree = doc.FirstChildElement(xmlutil::qn("p", "sld").c_str());
        if (!spTree) spTree = doc.RootElement();
        if (!spTree) return {};
        spTree = spTree->FirstChildElement(xmlutil::qn("p", "cSld").c_str());
        if (!spTree) return {};
        spTree = spTree->FirstChildElement(xmlutil::qn("p", "spTree").c_str());
        if (!spTree) return {};

        Slide slide;
        std::vector<RawConnector> raw;
        walk(spTree, slide.shapes, slide.tables, slide.pictures, raw);

        for (auto& rc : raw) resolveConnector(rc, slide.shapes, slide.edges);
        return slide;
    }

    // Find the shape whose center is closest to (px, py). Returns nullptr if
    // the connector's free end is unreasonably far from every shape (e.g. it
    // is anchored to the slide background, not to another node).
    static const Shape* nearestShape(const std::vector<Shape>& shapes, int64_t px, int64_t py) {
        const Shape* best = nullptr;
        int64_t bestDist = std::numeric_limits<int64_t>::max();
        for (auto& s : shapes) {
            int64_t cx = s.x + s.w / 2;
            int64_t cy = s.y + s.h / 2;
            int64_t dx = cx - px;
            int64_t dy = cy - py;
            int64_t d2 = dx * dx + dy * dy;
            if (d2 < bestDist) {
                bestDist = d2;
                best = &s;
            }
        }
        return best;
    }

    // Compute the (sx, sy) / (ex, ey) endpoints of a connector accounting for
    // the flipH / flipV attributes on its xfrm. Empty if xfrm cannot be parsed.
    static bool connectorEndpoints(const RawConnector& rc, int64_t& sx, int64_t& sy, int64_t& ex, int64_t& ey) {
        if (rc.w == 0 && rc.h == 0) return false;
        int64_t lx = rc.x, ly = rc.y, rx = rc.x + rc.w, ry = rc.y + rc.h;
        if (rc.flipH) std::swap(lx, rx);
        if (rc.flipV) std::swap(ly, ry);
        sx = lx; sy = ly; ex = rx; ey = ry;
        return true;
    }

    static void resolveConnector(const RawConnector& rc, const std::vector<Shape>& shapes, std::vector<Edge>& edges) {
        if (shapes.empty()) return;

        int64_t sx = 0, sy = 0, ex = 0, ey = 0;
        bool haveEnds = connectorEndpoints(rc, sx, sy, ex, ey);

        std::string from = rc.start_id;
        std::string to = rc.end_id;

        if (haveEnds) {
            // Reject implausible resolutions: a connector's free end must
            // land within ~2000000 EMU of a shape center, otherwise it's
            // likely anchored to the slide background.
            constexpr int64_t kMaxDist = 2000000LL * 2000000LL;

            if (from.empty()) {
                if (auto* s = nearestShape(shapes, sx, sy)) {
                    int64_t dx = (s->x + s->w / 2) - sx;
                    int64_t dy = (s->y + s->h / 2) - sy;
                    if (dx * dx + dy * dy <= kMaxDist) from = s->id;
                }
            }
            if (to.empty()) {
                if (auto* s = nearestShape(shapes, ex, ey)) {
                    int64_t dx = (s->x + s->w / 2) - ex;
                    int64_t dy = (s->y + s->h / 2) - ey;
                    if (dx * dx + dy * dy <= kMaxDist) to = s->id;
                }
            }
        }

        if (from.empty() || to.empty() || from == to) return;
        edges.push_back({from, to});
        if (rc.bidirectional) edges.push_back({to, from});
    }

private:
    void walk(const tinyxml2::XMLElement* parent,
              std::vector<Shape>& shapes,
              std::vector<Table>& tables,
              std::vector<Picture>& pictures,
              std::vector<RawConnector>& raws) {
        for (const tinyxml2::XMLElement* e = parent->FirstChildElement(); e;
             e = e->NextSiblingElement()) {
            std::string name = e->Name() ? e->Name() : "";
            // Strip "prefix:" if present
            auto local = [](const std::string& qn) {
                auto pos = qn.find(':');
                return pos == std::string::npos ? qn : qn.substr(pos + 1);
            };
            std::string tag = local(name);

            if (tag == "sp") {
                addShape(e, shapes);
            } else if (tag == "cxnSp") {
                addConnector(e, raws);
            } else if (tag == "grpSp") {
                // Recurse into group.
                walk(e, shapes, tables, pictures, raws);
            } else if (tag == "graphicFrame") {
                addTable(e, tables);
            } else if (tag == "pic") {
                addPicture(e, pictures);
            }
        }
    }

    void addShape(const tinyxml2::XMLElement* sp, std::vector<Shape>& shapes) {
        const tinyxml2::XMLElement* nvSpPr = sp->FirstChildElement(xmlutil::qn("p", "nvSpPr").c_str());
        const tinyxml2::XMLElement* cNvPr = nvSpPr ? nvSpPr->FirstChildElement(xmlutil::qn("p", "cNvPr").c_str()) : nullptr;
        if (!cNvPr) return;
        const char* idAttr = cNvPr->Attribute("id");
        if (!idAttr) return;
        Shape s;
        s.id = idAttr;
        if (const char* nm = cNvPr->Attribute("name")) s.name = nm;

        // Placeholder detection: <p:ph> can live either under <p:cNvSpPr> (legacy
        // location) or under <p:nvPr> (the current spec). Check both.
        auto readPh = [&](const tinyxml2::XMLElement* parent) -> const tinyxml2::XMLElement* {
            return parent ? parent->FirstChildElement(xmlutil::qn("p", "ph").c_str()) : nullptr;
        };
        const tinyxml2::XMLElement* cNvSpPr = nvSpPr->FirstChildElement(xmlutil::qn("p", "cNvSpPr").c_str());
        const tinyxml2::XMLElement* nvPr   = nvSpPr->FirstChildElement(xmlutil::qn("p", "nvPr").c_str());
        const tinyxml2::XMLElement* ph = readPh(cNvSpPr);
        if (!ph) ph = readPh(nvPr);
        if (ph) {
            s.phIdx = ph->IntAttribute("idx", -1);
            const char* typeAttr = ph->Attribute("type");
            if (typeAttr) {
                std::string t = typeAttr;
                if (t == "title" || t == "ctrTitle") s.isTitle = true;
                if (t == "body" || t == "subTitle") s.isBody = true;
            } else {
                // idx-only placeholder is typically a body
                if (s.phIdx >= 0) s.isBody = true;
            }
        }

        // txBody holds the rich-text body.
        const tinyxml2::XMLElement* txBody = sp->FirstChildElement(xmlutil::qn("p", "txBody").c_str());
        s.paragraphs = xmlutil::collectParagraphs(txBody);
        // Keep the joined text around for the Mermaid path / fallback use.
        for (size_t i = 0; i < s.paragraphs.size(); ++i) {
            if (i) s.text += '\n';
            s.text += s.paragraphs[i].second;
        }

        // Placeholder shapes inherit their position from the layout master;
        // accept them even without an explicit xfrm on the slide.
        bool isPlaceholder = s.isTitle || s.isBody || s.phIdx >= 0;
        if (!xmlutil::getXfrm(sp, s.x, s.y, s.w, s.h) && !isPlaceholder) return;
        shapes.push_back(std::move(s));
    }

    void addTable(const tinyxml2::XMLElement* gf, std::vector<Table>& tables) {
        const tinyxml2::XMLElement* nvGraphicFramePr = gf->FirstChildElement(xmlutil::qn("p", "nvGraphicFramePr").c_str());
        const tinyxml2::XMLElement* cNvPr = nvGraphicFramePr ? nvGraphicFramePr->FirstChildElement(xmlutil::qn("p", "cNvPr").c_str()) : nullptr;
        Table t;
        if (cNvPr && cNvPr->Attribute("name")) t.name = cNvPr->Attribute("name");

        // Drill into <a:graphic><a:graphicData><a:tbl>
        const tinyxml2::XMLElement* tbl = nullptr;
        for (const tinyxml2::XMLElement* a_g = gf->FirstChildElement(xmlutil::qn("a", "graphic").c_str()); a_g && !tbl;
             a_g = a_g->NextSiblingElement(xmlutil::qn("a", "graphic").c_str())) {
            const tinyxml2::XMLElement* a_gd = a_g->FirstChildElement(xmlutil::qn("a", "graphicData").c_str());
            for (const tinyxml2::XMLElement* c = a_gd ? a_gd->FirstChildElement() : nullptr; c && !tbl;
                 c = c->NextSiblingElement()) {
                std::string n = c->Name() ? c->Name() : "";
                auto colon = n.find(':');
                std::string local = colon == std::string::npos ? n : n.substr(colon + 1);
                if (local == "tbl") tbl = c;
            }
        }
        if (!tbl) return;

        for (const tinyxml2::XMLElement* tr = tbl->FirstChildElement(xmlutil::qn("a", "tr").c_str()); tr;
             tr = tr->NextSiblingElement(xmlutil::qn("a", "tr").c_str())) {
            TableRow row;
            for (const tinyxml2::XMLElement* tc = tr->FirstChildElement(xmlutil::qn("a", "tc").c_str()); tc;
                 tc = tc->NextSiblingElement(xmlutil::qn("a", "tc").c_str())) {
                const tinyxml2::XMLElement* txBody = tc->FirstChildElement(xmlutil::qn("a", "txBody").c_str());
                TableCell cell;
                cell.text = xmlutil::collectText(txBody);
                // Markdown tables don't preserve line breaks inside a cell;
                // collapse newlines to spaces.
                for (size_t i = 0; i < cell.text.size(); ++i) {
                    if (cell.text[i] == '\n') cell.text[i] = ' ';
                }
                row.cells.push_back(std::move(cell));
            }
            t.rows.push_back(std::move(row));
        }
        if (!t.rows.empty()) tables.push_back(std::move(t));
    }

    void addPicture(const tinyxml2::XMLElement* pic, std::vector<Picture>& pictures) {
        const tinyxml2::XMLElement* nvPicPr = pic->FirstChildElement(xmlutil::qn("p", "nvPicPr").c_str());
        const tinyxml2::XMLElement* cNvPr = nvPicPr ? nvPicPr->FirstChildElement(xmlutil::qn("p", "cNvPr").c_str()) : nullptr;
        if (!cNvPr) return;
        Picture p;
        if (auto* nm = cNvPr->Attribute("name")) p.name = nm;
        if (auto* d = cNvPr->Attribute("descr")) p.descr = d;
        const tinyxml2::XMLElement* blipFill = pic->FirstChildElement(xmlutil::qn("p", "blipFill").c_str());
        if (blipFill) {
            // blipFill/blip -> a:blip r:embed="rId..."
            for (const tinyxml2::XMLElement* blip = blipFill->FirstChildElement(xmlutil::qn("a", "blip").c_str()); blip;
                 blip = blip->NextSiblingElement(xmlutil::qn("a", "blip").c_str())) {
                if (auto* embed = blip->Attribute(xmlutil::qn("r", "embed").c_str())) {
                    p.relId = embed;
                    break;
                }
            }
        }
        pictures.push_back(std::move(p));
    }

    void addConnector(const tinyxml2::XMLElement* cxn, std::vector<RawConnector>& raws) {
        RawConnector rc;
        const tinyxml2::XMLElement* nvCxnSpPr = cxn->FirstChildElement(xmlutil::qn("p", "nvCxnSpPr").c_str());
        const tinyxml2::XMLElement* cNvCxnSpPr = nvCxnSpPr ? nvCxnSpPr->FirstChildElement(xmlutil::qn("p", "cNvCxnSpPr").c_str()) : nullptr;
        if (cNvCxnSpPr) {
            if (auto* st = cNvCxnSpPr->FirstChildElement(xmlutil::qn("a", "stCxn").c_str())) {
                if (auto* a = st->Attribute("id")) rc.start_id = a;
            }
            if (auto* en = cNvCxnSpPr->FirstChildElement(xmlutil::qn("a", "endCxn").c_str())) {
                if (auto* a = en->Attribute("id")) rc.end_id = a;
            }
        }
        if (rc.start_id.empty() && rc.end_id.empty()) return;

        // Read the connector's own bounding box (and flips) so we can resolve
        // a missing endpoint from its actual position on the slide.
        const tinyxml2::XMLElement* spPr = cxn->FirstChildElement(xmlutil::qn("p", "spPr").c_str());
        if (spPr) {
            const tinyxml2::XMLElement* xfrm = spPr->FirstChildElement(xmlutil::qn("a", "xfrm").c_str());
            if (xfrm) {
                if (auto* off = xfrm->FirstChildElement(xmlutil::qn("a", "off").c_str())) {
                    rc.x = off->Int64Attribute("x", 0);
                    rc.y = off->Int64Attribute("y", 0);
                }
                if (auto* ext = xfrm->FirstChildElement(xmlutil::qn("a", "ext").c_str())) {
                    rc.w = ext->Int64Attribute("cx", 0);
                    rc.h = ext->Int64Attribute("cy", 0);
                }
                rc.flipH = xfrm->BoolAttribute("flipH", false);
                rc.flipV = xfrm->BoolAttribute("flipV", false);
            }
            // Detect bidirectional arrowheads on the line.
            const tinyxml2::XMLElement* ln = spPr->FirstChildElement(xmlutil::qn("a", "ln").c_str());
            if (ln) {
                bool head = ln->FirstChildElement(xmlutil::qn("a", "headEnd").c_str()) != nullptr;
                bool tail = ln->FirstChildElement(xmlutil::qn("a", "tailEnd").c_str()) != nullptr;
                rc.bidirectional = head && tail;
            }
        }
        raws.push_back(std::move(rc));
    }
};

// ---------------------------------------------------------------------------
// SmartArt parser: reads the logical model from data*.xml.
// ---------------------------------------------------------------------------
class SmartArtParser {
public:
    // data_xml: contents of ppt/diagrams/dataN.xml
    // drawing_xml: contents of ppt/diagrams/drawingN.xml (used for fallback text)
    std::pair<std::vector<Shape>, std::vector<Edge>>
    parse(const std::string& data_xml, const std::string& drawing_xml = "") {
        std::vector<Shape> shapes;
        std::vector<Edge> edges;
        std::map<std::string, std::string> textByPtModelId;
        std::map<std::string, std::string> modelIdByPtModelId;

        tinyxml2::XMLDocument data;
        if (data.Parse(data_xml.c_str(), data_xml.size()) != tinyxml2::XML_SUCCESS) return {shapes, edges};

        // Walk <dgm:ptModel>.
        for (const tinyxml2::XMLElement* pt = data.FirstChildElement(xmlutil::qn("dgm", "ptLst").c_str())
                                            ? data.FirstChildElement(xmlutil::qn("dgm", "ptLst").c_str())->FirstChildElement(xmlutil::qn("dgm", "pt").c_str())
                                            : nullptr;
             pt; pt = pt->NextSiblingElement(xmlutil::qn("dgm", "pt").c_str())) {
            const char* modelId = pt->Attribute("modelId");
            if (!modelId) continue;
            const tinyxml2::XMLElement* prSet = pt->FirstChildElement(xmlutil::qn("dgm", "prSet").c_str());
            (void)prSet;
            const tinyxml2::XMLElement* t = pt->FirstChildElement(xmlutil::qn("dgm", "t").c_str());
            if (t && t->GetText()) textByPtModelId[modelId] = t->GetText();
            modelIdByPtModelId[modelId] = modelId;
        }

        // Walk <dgm:cxnLst><dgm:cxn>.
        if (const tinyxml2::XMLElement* cxnLst = data.FirstChildElement(xmlutil::qn("dgm", "cxnLst").c_str())) {
            for (const tinyxml2::XMLElement* cxn = cxnLst->FirstChildElement(xmlutil::qn("dgm", "cxn").c_str()); cxn;
                 cxn = cxn->NextSiblingElement(xmlutil::qn("dgm", "cxn").c_str())) {
                const tinyxml2::XMLElement* st = cxn->FirstChildElement(xmlutil::qn("dgm", "st").c_str());
                const tinyxml2::XMLElement* en = cxn->FirstChildElement(xmlutil::qn("dgm", "end").c_str());
                if (!st || !en) continue;
                const char* stId = st->Attribute("modelId");
                const char* enId = en->Attribute("modelId");
                if (!stId || !enId) continue;
                edges.push_back({stId, enId});
            }
        }

        // Build shapes from collected modelIds (dedup by modelId).
        std::set<std::string> seen;
        for (auto& [mid, _] : modelIdByPtModelId) {
            if (seen.count(mid)) continue;
            seen.insert(mid);
            Shape s;
            s.id = mid;
            s.name = "smartart:" + mid;
            auto it = textByPtModelId.find(mid);
            s.text = (it != textByPtModelId.end()) ? it->second : "";
            shapes.push_back(std::move(s));
        }

        // Fallback text from drawing if any model node lacked text.
        if (drawing_xml.empty() == false) {
            tinyxml2::XMLDocument drw;
            if (drw.Parse(drawing_xml.c_str(), drawing_xml.size()) == tinyxml2::XML_SUCCESS) {
                collectDrawingTexts(drw.RootElement(), textByPtModelId, shapes);
            }
        }

        return {shapes, edges};
    }

private:
    void collectDrawingTexts(const tinyxml2::XMLElement* root,
                             std::map<std::string, std::string>& /*textByPt*/,
                             std::vector<Shape>& shapes) {
        // Map every <a:t> found to the nearest enclosing shape's id. This is a
        // best-effort fallback when the logical model has empty text.
        for (const tinyxml2::XMLElement* sp = root; sp; sp = sp->NextSiblingElement()) {
            // recursive walk handled by collecting all <a:t> and matching
            // nearest <p:cNvPr id="..."> ancestor.
            walkForText(sp, shapes);
        }
    }

    void walkForText(const tinyxml2::XMLElement* e, std::vector<Shape>& shapes) {
        if (!e) return;
        std::string ename = e->Name() ? e->Name() : "";
        // Match any element whose local tag is "sp" (e.g. "p:sp" or "dgm:sp").
        auto colon = ename.find(':');
        std::string local = colon == std::string::npos ? ename : ename.substr(colon + 1);
        if (local == "sp") {
            const tinyxml2::XMLElement* cNvPr = e->FirstChildElement(xmlutil::qn("p", "nvSpPr").c_str());
            cNvPr = cNvPr ? cNvPr->FirstChildElement(xmlutil::qn("p", "cNvPr").c_str()) : nullptr;
            if (cNvPr) {
                const char* idAttr = cNvPr->Attribute("modelId");
                if (!idAttr) idAttr = cNvPr->Attribute("id");
                if (idAttr) {
                    std::string txt = xmlutil::collectText(e);
                    if (!txt.empty()) {
                        for (auto& s : shapes) {
                            if (s.id == idAttr && s.text.empty()) s.text = txt;
                        }
                    }
                }
            }
        }
        for (const tinyxml2::XMLElement* c = e->FirstChildElement(); c; c = c->NextSiblingElement()) {
            walkForText(c, shapes);
        }
    }
};

// ---------------------------------------------------------------------------
// Mermaid writer.
// ---------------------------------------------------------------------------
namespace mermaid {

std::string sanitizeId(const std::string& in) {
    std::string s;
    for (char c : in) {
        if (std::isalnum(static_cast<unsigned char>(c)) || c == '_') s += c;
        else s += '_';
    }
    if (s.empty()) s = "node";
    if (std::isdigit(static_cast<unsigned char>(s[0]))) s = "_" + s;
    return s;
}

// Escape for a quoted Mermaid label.
std::string escapeLabel(const std::string& in) {
    std::string out;
    out.reserve(in.size());
    for (char c : in) {
        switch (c) {
        case '"': out += "\\\""; break;
        case '\\': out += "\\\\"; break;
        case '\n': out += "<br/>"; break;
        case '\r': break;
        case '<': out += "&lt;"; break;
        case '>': out += "&gt;"; break;
        case '|': out += "&#124;"; break;
        default:
            if (static_cast<unsigned char>(c) < 0x20) {
                // skip control chars
            } else {
                out += c;
            }
        }
    }
    return out;
}

std::string detectDirection(const std::vector<Shape>& shapes, const std::vector<Edge>& edges) {
    int64_t dx = 0, dy = 0;
    int64_t cx = 0, cy = 0, wsum = 0, hsum = 0;
    for (auto& s : shapes) {
        cx += (s.x + s.w / 2) * s.h; cy += (s.y + s.h / 2) * s.w;
        wsum += s.w; hsum += s.h;
    }
    if (edges.empty()) return "TD";
    for (auto& e : edges) {
        const Shape* a = nullptr; const Shape* b = nullptr;
        for (auto& s : shapes) {
            if (s.id == e.from) a = &s;
            if (s.id == e.to)   b = &s;
        }
        if (!a || !b) continue;
        dx += std::abs((b->x + b->w / 2) - (a->x + a->w / 2));
        dy += std::abs((b->y + b->h / 2) - (a->y + a->h / 2));
    }
    return dx > dy ? "LR" : "TD";
}

std::string render(const std::vector<Shape>& shapes, const std::vector<Edge>& edges,
                   const std::string& direction, const std::string& title = "") {
    std::ostringstream os;
    if (!title.empty()) {
        os << "---\ntitle: " << title << "\n---\n";
    }
    os << "flowchart " << direction << "\n";

    // Deduplicate ids while keeping stable mermaid-safe identifiers.
    std::map<std::string, std::string> idMap;
    std::set<std::string> usedShortIds;
    for (auto& s : shapes) {
        std::string base = sanitizeId(s.id);
        std::string candidate = base;
        int n = 2;
        while (usedShortIds.count(candidate)) {
            candidate = base + "_" + std::to_string(n++);
        }
        usedShortIds.insert(candidate);
        idMap[s.id] = candidate;
    }

    for (auto& s : shapes) {
        std::string text = s.text.empty() ? s.name : s.text;
        std::string label = escapeLabel(text);
        os << "    " << idMap[s.id] << "[\"" << label << "\"]\n";
    }

    std::set<std::pair<std::string, std::string>> dedup;
    for (auto& e : edges) {
        auto itA = idMap.find(e.from);
        auto itB = idMap.find(e.to);
        if (itA == idMap.end() || itB == idMap.end()) continue;
        auto k = std::make_pair(itA->second, itB->second);
        if (dedup.count(k)) continue;
        dedup.insert(k);
        os << "    " << itA->second << " --> " << itB->second << "\n";
    }
    return os.str();
}

} // namespace mermaid

// ---------------------------------------------------------------------------
// Driver.
// ---------------------------------------------------------------------------
struct Options {
    std::string input;
    std::string output;
    std::vector<int> slides;          // 1-indexed, empty == all
    std::string direction = "auto";   // "TD" | "LR" | "auto"
    bool includeSmartArt = true;
    bool wrapInCodeFence = true;
    bool oneSectionPerSlide = true;
};

struct Section {
    std::string title;
    std::string body;
    bool isMermaid = false;
};

// Markdown renderer for slides without a box diagram. Emits titles as
// headings, plain text shapes as nested bullet lists whose indent matches
// each paragraph's <a:pPr lvl="N"/>, tables as Markdown tables, and pictures
// as image placeholders. Headings follow a 3-level hierarchy:
//   # document title  (H1, set elsewhere)
//   ## Slide N        (H2, set elsewhere)
//   ### Title         (H3, slide title placeholder)
//   ### content...
static std::string renderMarkdown(const Slide& slide) {
    std::ostringstream os;

    // 1. Title placeholder(s) come first.
    bool emittedAny = false;
    for (auto& s : slide.shapes) {
        if (s.isTitle && !s.text.empty()) {
            os << "# " << s.text << "\n\n";
            emittedAny = true;
        }
    }

    // 2. Tables - render before free text so they appear in document order.
    for (auto& t : slide.tables) {
        if (t.rows.empty()) continue;
        size_t ncols = 0;
        for (auto& r : t.rows) ncols = std::max(ncols, r.cells.size());
        if (ncols == 0) continue;
        // Header row uses the first row of the source table. If it looks
        // empty, fall back to a synthetic header.
        bool firstEmpty = true;
        for (auto& c : t.rows.front().cells) if (!c.text.empty()) firstEmpty = false;
        if (firstEmpty) {
            os << "| " << std::string(ncols, ' ') << " |" << std::string(ncols - 1, '|') << "\n";
        } else {
            for (size_t c = 0; c < ncols; ++c) {
                if (c) os << " | ";
                os << (c < t.rows.front().cells.size() ? t.rows.front().cells[c].text : "");
            }
            os << " |\n";
        }
        os << "| " << std::string(ncols, '-') << " |" << std::string(ncols - 1, '|') << "\n";
        size_t firstDataRow = firstEmpty ? 0 : 1;
        for (size_t r = firstDataRow; r < t.rows.size(); ++r) {
            for (size_t c = 0; c < ncols; ++c) {
                if (c) os << " | ";
                std::string cellText = c < t.rows[r].cells.size() ? t.rows[r].cells[c].text : "";
                // Escape pipe characters inside table cells.
                std::string esc;
                for (char ch : cellText) {
                    if (ch == '|') esc += "\\|";
                    else if (ch == '\r') continue;
                    else esc += ch;
                }
                os << esc;
            }
            os << " |\n";
        }
        os << "\n";
        emittedAny = true;
    }

    // 3. Pictures - reference rels that the caller can resolve later.
    for (auto& p : slide.pictures) {
        os << "![";
        os << (p.descr.empty() ? p.name : p.descr);
        os << "](media/" << (p.relId.empty() ? p.name : p.relId) << ")\n\n";
        emittedAny = true;
    }

    // 4. Text shapes (skip titles and shapes with no text). Render paragraphs
    //    as nested bullet lists - the level from <a:pPr lvl="N"/> drives
    //    indentation so progressive relationships stay visible.
    for (auto& s : slide.shapes) {
        if (s.isTitle) continue;
        if (s.text.empty()) continue;
        if (s.paragraphs.empty()) continue;
        bool wroteBullet = false;
        for (auto& [lvl, text] : s.paragraphs) {
            if (text.empty()) continue;
            // 2 spaces per level of indent.
            for (int i = 0; i < lvl; ++i) os << "  ";
            os << "- " << text << "\n";
            wroteBullet = true;
        }
        if (wroteBullet) {
            os << "\n";
            emittedAny = true;
        }
    }

    if (!emittedAny) return {};
    return os.str();
}

static std::vector<Section> convert(const Options& opts, PptxReader& z) {
    std::vector<Section> sections;
    auto allFiles = z.listFiles("ppt/slides/");
    std::vector<std::string> slideFiles;
    for (auto& f : allFiles) {
        if (f.rfind("ppt/slides/slide", 0) == 0 && f.size() > 4 && f.substr(f.size() - 4) == ".xml") {
            slideFiles.push_back(f);
        }
    }
    std::sort(slideFiles.begin(), slideFiles.end());

    SlideParser slideParser;
    SmartArtParser smartArtParser;

    for (size_t i = 0; i < slideFiles.size(); ++i) {
        if (!opts.slides.empty()) {
            int oneBased = static_cast<int>(i + 1);
            if (std::find(opts.slides.begin(), opts.slides.end(), oneBased) == opts.slides.end()) continue;
        }
        std::string xml = z.read(slideFiles[i]);
        if (xml.empty()) continue;
        Slide slide = slideParser.parse(xml);

        if (opts.includeSmartArt) {
            // Find diagrams referenced by this slide via the slide rels.
            std::string relsPath = slideFiles[i];
            relsPath.replace(relsPath.find("slides/"), 7, "slides/_rels/");
            relsPath += ".rels";
            std::string relsXml = z.read(relsPath);
            if (!relsXml.empty()) {
                tinyxml2::XMLDocument rels;
                if (rels.Parse(relsXml.c_str(), relsXml.size()) == tinyxml2::XML_SUCCESS) {
                    for (const tinyxml2::XMLElement* rel = rels.FirstChildElement(); rel;
                         rel = rel->NextSiblingElement()) {
                        const char* type = rel->Attribute("Type");
                        if (!type) continue;
                        std::string t = type;
                        if (t.find("/diagram") == std::string::npos) continue;
                        const char* target = rel->Attribute("Target");
                        if (!target) continue;
                        fs::path base = fs::path(slideFiles[i]).parent_path();
                        fs::path targetPath = base / target;
                        std::string normalized = targetPath.generic_string();
                        while (normalized.rfind("../", 0) == 0) normalized = normalized.substr(3);
                        std::string dataXml = z.read(normalized);
                        if (dataXml.empty()) continue;
                        std::string drawingXml;
                        fs::path p = normalized;
                        std::string drawingName = p.stem().string();
                        if (drawingName.rfind("data", 0) == 0) {
                            drawingName = "drawing" + drawingName.substr(4);
                            fs::path dp = p.parent_path() / (drawingName + p.extension().string());
                            drawingXml = z.read(dp.generic_string());
                        }
                        auto [sShapes, sEdges] = smartArtParser.parse(dataXml, drawingXml);
                        for (auto& s : sShapes) slide.shapes.push_back(std::move(s));
                        for (auto& e : sEdges) slide.edges.push_back(e);
                    }
                }
            }
        }

        Section sec;
        sec.title = "Slide " + std::to_string(i + 1);

        // Decide which output to emit.
        // Priority: if the slide has any connectors (or SmartArt edges), it's
        // a diagram - emit Mermaid. Otherwise emit Markdown text content.
        // A slide can have both: in that case emit Mermaid first then MD
        // for any leftover text shapes that aren't nodes of the diagram.
        bool hasDiagram = !slide.edges.empty();
        std::ostringstream out;
        if (hasDiagram) {
            std::string direction = opts.direction;
            if (direction == "auto") direction = mermaid::detectDirection(slide.shapes, slide.edges);
            std::string body = mermaid::render(slide.shapes, slide.edges, direction);
            if (opts.wrapInCodeFence) {
                out << "```mermaid\n" << body << "```\n";
            } else {
                out << body << "\n";
            }
            sec.isMermaid = true;

            // Emit any non-diagram text as Markdown below the chart.
            std::vector<Shape> leftover;
            std::set<std::string> edgeIds;
            for (auto& edge : slide.edges) { edgeIds.insert(edge.from); edgeIds.insert(edge.to); }
            for (auto& s : slide.shapes) {
                if (!edgeIds.count(s.id) && !s.text.empty() && !s.isTitle) leftover.push_back(s);
            }
            if (!leftover.empty()) {
                Slide fake;
                fake.shapes = std::move(leftover);
                fake.tables = slide.tables;
                fake.pictures = slide.pictures;
                std::string md = renderMarkdown(fake);
                if (!md.empty()) out << "\n" << md;
            }
        } else {
            std::string md = renderMarkdown(slide);
            if (md.empty()) continue;  // no extractable content - skip slide
            out << md;
        }

        sec.body = out.str();
        sections.push_back(std::move(sec));
    }
    return sections;
}

static int usage() {
    std::cerr <<
R"(pptx_to_mermaid - convert PPTX box diagrams to Mermaid flowcharts.

Usage:
  pptx_to_mermaid <input.pptx> [output.md] [options]

Options:
  -s, --slides N [N ...]   Only convert the listed slides (1-based).
  -d, --direction DIR      Flowchart direction: TD | LR | BT | RL (default: auto).
      --no-smartart        Skip SmartArt diagrams on slides.
      --no-fence           Emit raw Mermaid without ``` fences.
      --one-section        Emit one ## heading + fenced block per slide (default).
      --single             Emit a single fenced block (combine all slides with ----).
  -h, --help               Show this help.

Examples:
  pptx_to_mermaid deck.pptx
  pptx_to_mermaid deck.pptx out.md --slides 1 3
  pptx_to_mermaid deck.pptx -d LR
)";
    return 1;
}

int main(int argc, char** argv) {
    Options opts;
    std::vector<std::string> args;
    for (int i = 1; i < argc; ++i) args.emplace_back(argv[i]);
    if (args.empty()) return usage();

    // Pass 1: flags. Pass 2: positional input/output.
    size_t i = 0;
    auto needArg = [&](const std::string& flag) -> const std::string* {
        if (i + 1 >= args.size()) {
            std::cerr << "Missing value for " << flag << "\n";
            return nullptr;
        }
        return &args[++i];
    };
    while (i < args.size()) {
        const std::string& a = args[i];
        if (a == "-h" || a == "--help") { usage(); std::exit(0); }
        if (a == "--no-smartart") { opts.includeSmartArt = false; ++i; continue; }
        if (a == "--no-fence")    { opts.wrapInCodeFence = false; ++i; continue; }
        if (a == "--one-section") { opts.oneSectionPerSlide = true;  ++i; continue; }
        if (a == "--single")      { opts.oneSectionPerSlide = false; ++i; continue; }
        if (a == "-d" || a == "--direction") {
            const std::string* v = needArg(a); if (!v) return usage();
            opts.direction = *v; ++i; continue;
        }
        if (a == "-s" || a == "--slides") {
            ++i;
            while (i < args.size() && !args[i].empty() && args[i][0] != '-') {
                try { opts.slides.push_back(std::stoi(args[i])); }
                catch (...) { std::cerr << "Bad slide number: " << args[i] << "\n"; return usage(); }
                ++i;
            }
            continue;
        }
        if (!a.empty() && a[0] != '-') break; // first positional -> input
        std::cerr << "Unknown argument: " << a << "\n";
        return usage();
    }
    if (i >= args.size()) return usage();
    opts.input = args[i++];
    if (i < args.size() && !args[i].empty() && args[i][0] != '-') {
        opts.output = args[i++];
    }
    if (i < args.size()) {
        std::cerr << "Unexpected argument: " << args[i] << "\n";
        return usage();
    }

    if (opts.input.empty()) return usage();

    PptxReader z;
    if (!z.open(opts.input)) {
        std::cerr << "Failed to open: " << opts.input << "\n";
        return 2;
    }
    auto sections = convert(opts, z);
    z.close();

    if (sections.empty()) {
        std::cerr << "No box diagrams found in: " << opts.input << "\n";
        return 3;
    }

    std::ostringstream out;
    if (opts.oneSectionPerSlide) {
        for (auto& s : sections) {
            out << "## " << s.title << "\n\n" << s.body << "\n";
        }
    } else {
        out << "```mermaid\n";
        for (size_t i = 0; i < sections.size(); ++i) {
            if (i) out << "\n";
            // Strip the outer fence from each section so we can merge.
            std::string m = sections[i].body;
            auto trim = [&](const std::string& sub) {
                auto p = m.find(sub);
                if (p != std::string::npos) m.erase(p, sub.size());
            };
            trim("```mermaid\n");
            trim("```\n");
            out << m;
        }
        out << "```\n";
    }

    std::string result = out.str();
    if (opts.output.empty()) {
        std::cout << result;
    } else {
        std::ofstream f(opts.output, std::ios::binary);
        if (!f) {
            std::cerr << "Failed to write: " << opts.output << "\n";
            return 4;
        }
        f.write(result.data(), static_cast<std::streamsize>(result.size()));
        std::cerr << "Wrote " << result.size() << " bytes to " << opts.output << "\n";
    }
    return 0;
}
