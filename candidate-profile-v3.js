/* HANZ Universal Candidate Profile V3
 * General-purpose CV normalization and parser fallback.
 * No candidate-specific rules. Unknown fields are evidence gaps, not rejection signals.
 */

(() => {
  "use strict";

  const VERSION = "3.0.0";
  const originalAnalyzeCv = analyzeCv;
  const originalInferJobFamilies = inferJobFamilies;
  const originalScoreJob = scoreJob;
  const originalRenderAnalyzedProfile = renderAnalyzedProfile;

  const uniq = values => [...new Set((values || []).filter(Boolean).map(v => String(v).trim()).filter(Boolean))];
  const known = value => {
    if (value === null || value === undefined) return false;
    if (Array.isArray(value)) return value.length > 0;
    const s = String(value).trim();
    return !!s && s.toUpperCase() !== "UNKNOWN" && s.toLowerCase() !== "belum terdeteksi";
  };

  function loadScript(src, globalName){
    if (globalName && window[globalName]) return Promise.resolve(window[globalName]);
    return new Promise((resolve, reject) => {
      const existing = [...document.scripts].find(s => s.src === src);
      if (existing){
        if (globalName && window[globalName]) return resolve(window[globalName]);
        existing.addEventListener("load", () => resolve(globalName ? window[globalName] : true), {once:true});
        existing.addEventListener("error", reject, {once:true});
        return;
      }
      const s = document.createElement("script");
      s.src = src;
      s.async = true;
      s.onload = () => resolve(globalName ? window[globalName] : true);
      s.onerror = () => reject(new Error("CV extraction helper could not be loaded"));
      document.head.appendChild(s);
    });
  }

  async function extractCvText(file){
    if (!file) return {text:"", ok:false, source:"NONE"};
    const name = String(file.name || "").toLowerCase();
    const type = String(file.type || "").toLowerCase();

    if (name.endsWith(".txt") || type.includes("text/plain")){
      const text = await file.text();
      return {text, ok:!!text.trim(), source:"CLIENT_TXT"};
    }

    if (name.endsWith(".docx") || type.includes("wordprocessingml")){
      await loadScript("https://cdn.jsdelivr.net/npm/mammoth@1.8.0/mammoth.browser.min.js", "mammoth");
      const arrayBuffer = await file.arrayBuffer();
      const result = await window.mammoth.extractRawText({arrayBuffer});
      const text = String(result?.value || "");
      return {text, ok:!!text.trim(), source:"CLIENT_DOCX_MAMMOTH"};
    }

    if (name.endsWith(".pdf") || type.includes("pdf")){
      await loadScript("https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js", "pdfjsLib");
      window.pdfjsLib.GlobalWorkerOptions.workerSrc = "https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js";
      const data = new Uint8Array(await file.arrayBuffer());
      const pdf = await window.pdfjsLib.getDocument({data}).promise;
      const pages = [];
      for (let i=1; i<=pdf.numPages; i++){
        const page = await pdf.getPage(i);
        const content = await page.getTextContent();
        pages.push(content.items.map(x => x.str || "").join(" "));
      }
      const text = pages.join("\n");
      return {text, ok:!!text.trim(), source:"CLIENT_PDF_PDFJS"};
    }

    return {text:"", ok:false, source:"UNSUPPORTED_CLIENT_FORMAT"};
  }

  function collectServerText(result){
    const candidates = [
      result?.text,
      result?.raw_text,
      result?.extracted_text,
      result?.normalized_text,
      result?.candidate_profile?.text,
      result?.candidate_profile?.raw_text,
      result?.candidate_profile?.extracted_text,
      result?.candidate_profile?.normalized_text,
      result?.extraction?.text,
      result?.extraction?.raw_text,
      result?.extraction?.normalized_text
    ];
    return candidates.find(x => typeof x === "string" && x.trim()) || "";
  }

  function detectEducation(text){
    const h = normalize(text || "");
    if (/\b(phd|doctor|doctoral|doktor|s3)\b/.test(h)) return "S3";
    if (/\b(master|masters|magister|s2)\b/.test(h)) return "S2";
    if (/\b(bachelor|sarjana|s1|d4|diploma 4|bachelor_or_d4)\b/.test(h)) return "D4_S1";
    if (/\b(d3|diploma 3)\b/.test(h)) return "D3";
    if (/\b(sma|smk|high school)\b/.test(h)) return "SMA_SMK";
    return "UNKNOWN";
  }

  function detectGraduationYear(text){
    const h = String(text || "");
    const educationLines = h.split(/\r?\n/).filter(line => /education|pendidikan|university|universitas|politeknik|college|bachelor|sarjana|d4|s1/i.test(line));
    const years = [...educationLines.join(" ").matchAll(/\b(20\d{2}|19\d{2})\b/g)].map(m => Number(m[1]));
    if (!years.length) return null;
    return Math.max(...years.filter(y => y >= 1950 && y <= 2100));
  }

  function detectStudyPrograms(text){
    const h = normalize(text || "");
    const definitions = [
      ["Teknologi Rekayasa Internet", ["teknologi rekayasa internet", "internet engineering"]],
      ["Network Engineering", ["network engineering"]],
      ["Computer Science", ["computer science", "ilmu komputer"]],
      ["Informatics", ["informatika", "informatics"]],
      ["Information Technology", ["information technology", "teknologi informasi"]],
      ["Software Engineering", ["software engineering", "rekayasa perangkat lunak"]],
      ["Cybersecurity", ["cybersecurity", "cyber security", "keamanan siber"]],
      ["Telecommunication Engineering", ["telecommunication engineering", "teknik telekomunikasi"]],
      ["Mechatronics", ["mechatronics", "mekatronika"]],
      ["Electrical Engineering", ["electrical engineering", "teknik elektro"]],
      ["Chemical Engineering", ["chemical engineering", "teknik kimia"]],
      ["Industrial Engineering", ["industrial engineering", "teknik industri"]],
      ["International Business Management", ["international business management", "manajemen bisnis internasional"]],
      ["Business Management", ["business management", "manajemen bisnis"]],
      ["Business Administration", ["business administration", "administrasi bisnis"]],
      ["Marketing", ["marketing management", "manajemen pemasaran"]],
      ["Finance", ["finance", "keuangan"]],
      ["Accounting", ["accounting", "akuntansi"]],
      ["Economics", ["economics", "ekonomi"]]
    ];
    return definitions.filter(([,aliases]) => aliases.some(a => h.includes(normalize(a)))).map(([name]) => name).slice(0,5);
  }

  function detectSkills(text){
    const h = normalize(text || "");
    const defs = [
      ["Microsoft Excel", ["microsoft excel", "excel"]],
      ["Microsoft Office", ["microsoft office"]],
      ["PowerPoint", ["powerpoint"]],
      ["Marketing", ["marketing", "pemasaran"]],
      ["Digital Marketing", ["digital marketing", "pemasaran digital"]],
      ["Social Media Marketing", ["social media marketing", "media sosial"]],
      ["E-Commerce Marketing", ["e-commerce marketing", "ecommerce marketing", "tokopedia", "shopee"]],
      ["Brand Management", ["brand management", "brand awareness"]],
      ["Brand Promotion", ["brand promotion", "promosi merek"]],
      ["Campaign Execution", ["campaign execution", "campaign", "kampanye"]],
      ["Customer Acquisition", ["customer acquisition", "akuisisi pelanggan"]],
      ["Donor Acquisition", ["donor acquisition", "akuisisi donatur"]],
      ["Fundraising", ["fundraising", "fundraiser"]],
      ["Sales", ["sales", "penjualan"]],
      ["Business Development", ["business development", "pengembangan bisnis"]],
      ["Customer Retention", ["customer retention", "retensi pelanggan"]],
      ["Relationship Management", ["relationship management", "customer relationship", "client relationship"]],
      ["Market Research", ["market research", "riset pasar"]],
      ["Content Marketing", ["content marketing"]],
      ["Negotiation", ["negotiation", "negosiasi"]],
      ["Stakeholder Management", ["stakeholder management"]],
      ["Python", ["python"]],
      ["C++", ["c++"]],
      ["C", [" c "]],
      ["JavaScript", ["javascript"]],
      ["PHP", ["php"]],
      ["Laravel", ["laravel"]],
      ["React", ["react"]],
      ["MySQL", ["mysql"]],
      ["Firebase", ["firebase"]],
      ["Docker", ["docker"]],
      ["Kubernetes", ["kubernetes"]],
      ["Terraform", ["terraform"]],
      ["Ansible", ["ansible"]],
      ["AWS", ["aws", "amazon web services"]],
      ["Azure", ["azure"]],
      ["Google Cloud", ["google cloud", "gcp"]],
      ["TCP/IP", ["tcp/ip", "tcp ip"]],
      ["Networking", ["networking", "computer network", "jaringan komputer"]],
      ["Routing", ["routing"]],
      ["Switching", ["switching"]],
      ["GNS3", ["gns3"]],
      ["MikroTik", ["mikrotik"]],
      ["Cisco", ["cisco"]],
      ["BGP", ["bgp"]],
      ["OSPF", ["ospf"]],
      ["Cybersecurity", ["cybersecurity", "cyber security"]],
      ["SIEM", ["siem"]],
      ["Firewall", ["firewall"]],
      ["Linux", ["linux"]],
      ["IoT", ["iot", "internet of things"]],
      ["Arduino", ["arduino"]],
      ["Figma", ["figma"]],
      ["Git", [" github ", " git "]],
      ["PLC", ["plc"]],
      ["SCADA", ["scada"]],
      ["TIA Portal", ["tia portal"]],
      ["WinCC", ["wincc"]]
    ];
    return defs.filter(([,aliases]) => aliases.some(a => h.includes(normalize(a)))).map(([name]) => name).slice(0,35);
  }

  function detectLanguages(text){
    const h = normalize(text || "");
    const defs = [
      ["English", ["english", "bahasa inggris"]],
      ["Indonesian", ["indonesian", "bahasa indonesia"]],
      ["Japanese", ["japanese", "bahasa jepang"]],
      ["Arabic", ["arabic", "bahasa arab"]],
      ["Mandarin", ["mandarin", "chinese", "bahasa mandarin"]],
      ["Korean", ["korean", "bahasa korea"]]
    ];
    return defs.filter(([,aliases]) => aliases.some(a => h.includes(normalize(a)))).map(([name]) => name);
  }

  function detectCareerStatus(text, manual){
    if (known(manual?.career_status)) return manual.career_status;
    const h = normalize(text || "");
    if (/\bfresh graduate\b|\bnew graduate\b|\brecent graduate\b/.test(h)) return "FRESH_GRADUATE";
    if (/\bstudent\b|\bmahasiswa\b/.test(h) && !/\bwork experience\b|\bpengalaman kerja\b/.test(h)) return "STUDENT_OR_INTERNSHIP";
    if (/\bpresent\b|\bsekarang\b|\bsenior\b|\bmanager\b|\bofficer\b|\bfundraiser\b|\bbrand ambassador\b/.test(h)) return "EXPERIENCED";
    return "UNKNOWN";
  }

  function detectExperienceYears(text, manual){
    if (manual?.experience_years !== null && manual?.experience_years !== undefined) return manual.experience_years;
    const h = normalize(text || "");
    const m = h.match(/(?:over |more than |lebih dari )?(\d{1,2})\+?\s*(?:years?|yrs?|tahun)\s+(?:of )?(?:work )?(?:experience|pengalaman)/i);
    return m ? Number(m[1]) : null;
  }

  function businessFamilies(text){
    const h = normalize(text || "");
    const families = [];
    const has = terms => terms.some(t => h.includes(normalize(t)));
    if (has(["marketing", "pemasaran", "digital marketing", "brand awareness", "brand promotion", "campaign", "kampanye", "content marketing"])) families.push("Marketing / Growth");
    if (has(["sales", "penjualan", "business development", "pengembangan bisnis", "customer acquisition", "client acquisition", "donor acquisition", "fundraising", "fundraiser", "account management"])) families.push("Sales / Business Development");
    if (has(["international business management", "manajemen bisnis internasional", "business management", "business administration", "management", "manajemen", "business strategy"])) families.push("Business / Management");
    if (has(["customer success", "relationship management", "customer engagement", "donor engagement", "partnership", "stakeholder management", "customer retention", "retensi pelanggan"])) families.push("Customer Success / Partnerships");
    return families;
  }

  inferJobFamilies = function(studyPrograms, skills, extraText=""){
    let base = [];
    try { base = originalInferJobFamilies(studyPrograms, skills) || []; } catch {}
    const combined = [...(studyPrograms || []), ...(skills || []), extraText || ""].join(" ");
    return uniq([...base, ...businessFamilies(combined)]);
  };

  function buildFallbackProfile(text, manual, extraction){
    const studies = uniq([...(manual?.study_program ? [manual.study_program] : []), ...detectStudyPrograms(text)]);
    const skills = uniq([...(manual?.skills || []), ...detectSkills(text)]);
    return {
      education_level: known(manual?.education_level) ? manual.education_level : detectEducation(text),
      study_programs: studies,
      skills,
      career_status: detectCareerStatus(text, manual || {}),
      experience_years: detectExperienceYears(text, manual || {}),
      languages: detectLanguages(text),
      job_families: inferJobFamilies(studies, skills, text),
      nationality: manual?.nationality || "",
      graduation_year: manual?.graduation_year ?? detectGraduationYear(text),
      _cv_text_available: !!String(text || "").trim(),
      _extraction_source: extraction?.source || "UNKNOWN"
    };
  }

  function normalizeLanguages(values){
    return uniq((values || []).map(v => typeof v === "string" ? v : (v?.language || v?.name || v?.label || "")));
  }

  function mergeProfiles(server, fallback, manual){
    server = server || {};
    fallback = fallback || {};
    manual = manual || {};

    const merged = {
      ...server,
      education_level: known(manual.education_level) ? manual.education_level : known(server.education_level) ? server.education_level : known(fallback.education_level) ? fallback.education_level : "UNKNOWN",
      study_programs: uniq([...(manual.study_program ? [manual.study_program] : []), ...(server.study_programs || []), ...(fallback.study_programs || [])]),
      skills: uniq([...(manual.skills || []), ...(server.skills || []), ...(fallback.skills || [])]),
      career_status: known(manual.career_status) ? manual.career_status : known(server.career_status) ? server.career_status : known(fallback.career_status) ? fallback.career_status : "UNKNOWN",
      experience_years: manual.experience_years !== null && manual.experience_years !== undefined ? manual.experience_years : server.experience_years !== null && server.experience_years !== undefined ? server.experience_years : fallback.experience_years ?? null,
      languages: uniq([...normalizeLanguages(server.languages), ...(fallback.languages || [])]),
      nationality: manual.nationality || server.nationality || fallback.nationality || "",
      graduation_year: manual.graduation_year ?? server.graduation_year ?? fallback.graduation_year ?? null,
      preferred_work_mode: manual.preferred_work_mode ?? server.preferred_work_mode ?? "",
      target_countries: uniq(manual.target_countries?.length ? manual.target_countries : (server.target_countries || []))
    };

    merged.job_families = uniq([
      ...(server.job_families || []),
      ...(fallback.job_families || []),
      ...inferJobFamilies(merged.study_programs, merged.skills, [ ...(server.job_families || []), ...(fallback.job_families || []) ].join(" "))
    ]);

    const fields = {
      education: known(merged.education_level),
      study_program: merged.study_programs.length > 0,
      career_status: known(merged.career_status),
      experience: merged.experience_years !== null && merged.experience_years !== undefined,
      job_family: merged.job_families.length > 0,
      skills: merged.skills.length > 0,
      languages: merged.languages.length > 0
    };
    const total = Object.keys(fields).length;
    const found = Object.values(fields).filter(Boolean).length;
    merged.profile_health = {
      version: VERSION,
      completeness: Math.round((found / total) * 100),
      fields,
      extraction_source: fallback._extraction_source || "SERVER_ANALYZER",
      unknown_is_non_blocking: true
    };
    return merged;
  }

  analyzeCv = async function(file, manualProfile){
    let extraction = {text:"", ok:false, source:"CLIENT_EXTRACTION_FAILED"};
    try { extraction = await extractCvText(file); }
    catch (error){ console.warn("HANZ client extraction fallback unavailable", error); }

    let serverResult = null;
    let serverError = null;
    try { serverResult = await originalAnalyzeCv(file, manualProfile); }
    catch (error){ serverError = error; console.warn("HANZ server analyzer failed; attempting client fallback", error); }

    const serverText = collectServerText(serverResult);
    const rawText = extraction.text || serverText;
    if (!serverResult && !String(rawText || "").trim()) throw serverError || new Error("CV tidak dapat diekstrak oleh server maupun browser.");

    const fallback = buildFallbackProfile(rawText, manualProfile || {}, extraction);
    const merged = mergeProfiles(serverResult?.candidate_profile || {}, fallback, manualProfile || {});
    merged.profile_health.server_analyzer_ok = !!serverResult;
    merged.profile_health.client_extraction_ok = !!extraction.ok;
    merged.profile_health.client_extraction_source = extraction.source;

    return { ...(serverResult || {}), ok:true, candidate_profile:merged, analyzer_mode:"UNIVERSAL_MERGED_PROFILE_V3" };
  };

  scoreJob = function(job){
    const result = originalScoreJob(job);
    if (!candidateProfile || result?.hard_mismatch) return result;

    const jobText = normalize([job?.job_title, job?.company_name, ...(job?.study_fields || [])].join(" "));
    const terms = {
      "Marketing / Growth":["marketing", "digital marketing", "brand", "campaign", "growth", "customer acquisition", "market research", "content", "social media"],
      "Sales / Business Development":["sales", "business development", "account executive", "account manager", "partnership", "client acquisition", "customer acquisition", "fundraising", "fundraiser"],
      "Business / Management":["business", "management", "management trainee", "commercial", "strategy", "operations management"],
      "Customer Success / Partnerships":["customer success", "relationship management", "partnership", "stakeholder", "customer engagement", "client relationship"]
    };

    const family = (candidateProfile.job_families || []).find(f => (terms[f] || []).some(term => jobText.includes(normalize(term))));
    if (family){
      return {
        ...result,
        score:Math.min(100, Number(result?.score || 0) + 30),
        reasons:uniq([...(result?.reasons || []), "Job family sesuai dengan background: " + family]),
        eligible:true
      };
    }

    // Incomplete parser evidence must not become a silent rejection.
    const health = candidateProfile.profile_health;
    const incomplete = health && Number(health.completeness || 0) < 70;
    if (incomplete && result && !result.hard_mismatch && Number(result.score || 0) >= 25){
      return {...result, eligible:true, reasons:uniq([...(result.reasons || []), "Profil belum lengkap; lowongan dipertahankan untuk verifikasi, bukan ditolak otomatis."])};
    }
    return result;
  };

  renderAnalyzedProfile = function(){
    originalRenderAnalyzedProfile();
    const grid = document.getElementById("analysisGrid");
    const health = candidateProfile?.profile_health;
    if (!grid || !health) return;
    document.getElementById("hanzProfileHealth")?.remove();
    const missing = Object.entries(health.fields || {}).filter(([,ok]) => !ok).map(([key]) => key.replaceAll("_", " "));
    const div = document.createElement("div");
    div.id = "hanzProfileHealth";
    div.className = "analysis-item";
    div.innerHTML = `<strong>Parser Health</strong>${esc(health.completeness)}% complete<br>${missing.length ? `<span style="color:var(--orange)">Missing: ${esc(missing.join(", "))} — treated as unknown, not rejected.</span>` : `<span style="color:var(--green)">Profile extraction complete.</span>`}`;
    grid.appendChild(div);
  };

  window.HANZProfileV3 = {
    version:VERSION,
    extractCvText,
    buildFallbackProfile,
    mergeProfiles,
    selfTest(){
      const marketing = buildFallbackProfile("Sarjana S1 Manajemen Bisnis Internasional. Marketing, customer acquisition, fundraising, campaign execution, business development, sales.", {}, {source:"SELF_TEST"});
      const technical = buildFallbackProfile("Fresh graduate D4 Teknologi Rekayasa Internet. Python Docker TCP/IP Networking GNS3 IoT Arduino Laravel MySQL.", {}, {source:"SELF_TEST"});
      const marketingPass = marketing.job_families.includes("Marketing / Growth") && marketing.job_families.includes("Sales / Business Development");
      const technicalPass = technical.job_families.includes("Network / Infrastructure") && technical.job_families.includes("IoT / Embedded Systems");
      return {ok:marketingPass && technicalPass, marketing, technical};
    }
  };

  const test = window.HANZProfileV3.selfTest();
  if (!test.ok) console.error("HANZ Profile V3 regression self-test failed", test);
  else console.info("HANZ Profile V3 ready", test);
})();
