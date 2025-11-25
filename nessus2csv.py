
#!/usr/bin/env python3
import argparse
import csv
import os
import sys
import xml.etree.ElementTree as etree
import html

def parse_nessus_file(file_path):
    """
    Parse a .nessus XML file and return a list of dict rows with both
    vulnerability and compliance fields (if present).
    """
    rows = []
    try:
        tree = etree.parse(file_path)
        root = tree.getroot()
    except Exception as e:
        print(f"[!] Failed to parse '{file_path}': {e}", file=sys.stderr)
        return rows

    # Detect namespaces (especially the 'cm' compliance namespace)
    # Example Nessus exports use cm="http://www.nessus.org/cm"
    nsmap = {}
    for elem in root.iter():
        if elem.tag[0] == "{":
            uri, _, tag = elem.tag[1:].partition("}")
            # Try to discover a prefix by checking child tags with prefixes
            # If the file explicitly defines 'cm', we can safely set it
            # Fallback to common CM namespace URI
            if "nessus.org/cm" in uri or "compliance" in uri:
                nsmap["cm"] = uri
            # We only need cm; other namespaces are usually not necessary here
    # Fallback CM namespace URI commonly used in Nessus exports
    nsmap.setdefault("cm", "http://www.nessus.org/cm")

    severity_map = {
        "0": "Info",
        "1": "Low",
        "2": "Medium",
        "3": "High",
        "4": "Critical"
    }

    report_hosts = root.findall(".//ReportHost")
    for host in report_hosts:
        ip = host.get("name") or ""
        if not ip:
            # Fallback to HostProperties (host-ip tag)
            for tag in host.findall(".//tag"):
                if tag.get("name") == "host-ip" and tag.text:
                    ip = tag.text
                    break

        for item in host.findall(".//ReportItem"):
            # Base fields
            plugin_name = item.get("pluginName") or item.findtext("plugin_name") or ""
            plugin_id = item.get("pluginID") or ""
            plugin_family = item.get("pluginFamily") or ""
            severity_code = item.get("severity") or ""
            severity_text = severity_map.get(severity_code, severity_code if severity_code else "Unknown")
            port = item.get("port") or ""
            protocol = item.get("protocol") or ""
            svc_name = item.get("svc_name") or ""
            risk_factor = (item.findtext("risk_factor") or "").strip()

            # Compliance fields (namespace-aware)
            cm = nsmap["cm"]
            def cm_text(tagname):
                # Try ns-aware first; if not found, try literal tag (in case namespace missing)
                text = item.findtext(f"{{{cm}}}{tagname}")
                if text is None:
                    text = item.findtext(f"{tagname}")  # sometimes exported without namespace
                if text is None:
                    # Also check inside child structure (some fields sit under nested nodes)
                    elem = item.find(f".//{{{cm}}}{tagname}")
                    if elem is not None and elem.text:
                        text = elem.text
                return (html.unescape(text.strip()) if text else "")

            compliance = (item.findtext("compliance") or "").strip().lower() == "true"
            compliance_check_name = cm_text("compliance-check-name")
            compliance_check_id = cm_text("compliance-check-id")
            compliance_result = cm_text("compliance-result")
            compliance_policy_value = cm_text("compliance-policy-value")
            compliance_actual_value = cm_text("compliance-actual-value")
            compliance_solution = cm_text("compliance-solution")
            compliance_info = cm_text("compliance-info")
            compliance_reference = cm_text("compliance-reference")
            compliance_benchmark_name = cm_text("compliance-benchmark-name")
            compliance_benchmark_version = cm_text("compliance-benchmark-version")
            compliance_benchmark_profile = cm_text("compliance-benchmark-profile")
            compliance_full_id = cm_text("compliance-full-id")
            compliance_control_id = cm_text("compliance-control-id")
            compliance_see_also = cm_text("compliance-see-also")
            compliance_source = cm_text("compliance-source")
            compliance_audit_file = cm_text("compliance-audit-file")
            compliance_functional_id = cm_text("compliance-functional-id")
            compliance_informational_id = cm_text("compliance-informational-id")

            row = {
                "IP Address": ip,
                "Port": port,
                "Protocol": protocol,
                "Service": svc_name,
                "Plugin ID": plugin_id,
                "Plugin Name": plugin_name,
                "Plugin Family": plugin_family,
                "Severity": severity_text,
                "Risk Factor": risk_factor,
                "Is Compliance Item": "true" if compliance else "false",
                "Compliance Result": compliance_result,
                "Compliance Check Name": compliance_check_name,
                "Compliance Check ID": compliance_check_id,
                "Compliance Policy Value": compliance_policy_value,
                "Compliance Actual Value": compliance_actual_value,
                "Compliance Solution": compliance_solution,
                "Compliance Info": compliance_info,
                "Compliance Reference": compliance_reference,
                "Compliance Benchmark Name": compliance_benchmark_name,
                "Compliance Benchmark Version": compliance_benchmark_version,
                "Compliance Benchmark Profile": compliance_benchmark_profile,
                "Compliance Full ID": compliance_full_id,
                "Compliance Control ID": compliance_control_id,
                "Compliance See Also": compliance_see_also,
                "Compliance Source": compliance_source,
                "Compliance Audit File": compliance_audit_file,
                "Compliance Functional ID": compliance_functional_id,
                "Compliance Informational ID": compliance_informational_id,
            }
            rows.append(row)

    # If nothing parsed (strange structure), try a flat fallback
    if not rows:
        ips = [h.get("name") or "" for h in root.iter("ReportHost")]
        items = list(root.iter("ReportItem"))
        for idx, it in enumerate(items):
            plugin_name = it.get("pluginName") or it.findtext("plugin_name") or ""
            plugin_id = it.get("pluginID") or ""
            severity_code = it.get("severity") or ""
            severity_text = severity_map.get(severity_code, severity_code if severity_code else "Unknown")
            ip = ips[idx] if idx < len(ips) else ""
            rows.append({
                "IP Address": ip,
                "Port": it.get("port") or "",
                "Protocol": it.get("protocol") or "",
                "Service": it.get("svc_name") or "",
                "Plugin ID": plugin_id,
                "Plugin Name": plugin_name,
                "Plugin Family": it.get("pluginFamily") or "",
                "Severity": severity_text,
                "Risk Factor": (it.findtext("risk_factor") or "").strip(),
                "Is Compliance Item": "false",
                "Compliance Result": "",
                "Compliance Check Name": "",
                "Compliance Check ID": "",
                "Compliance Policy Value": "",
                "Compliance Actual Value": "",
                "Compliance Solution": "",
                "Compliance Info": "",
                "Compliance Reference": "",
                "Compliance Benchmark Name": "",
                "Compliance Benchmark Version": "",
                "Compliance Benchmark Profile": "",
                "Compliance Full ID": "",
                "Compliance Control ID": "",
                "Compliance See Also": "",
                "Compliance Source": "",
                "Compliance Audit File": "",
                "Compliance Functional ID": "",
                "Compliance Informational ID": "",
            })

    return rows


def collect_nessus_inputs(input_path, recursive=False):
    """
    Return a list of .nessus file paths from either a single file or directory.
    """
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input path does not exist: {input_path}")

    if os.path.isfile(input_path):
        if input_path.lower().endswith(".nessus"):
            return [input_path]
        else:
            raise ValueError(f"File is not a .nessus file: {input_path}")

    nessus_files = []
    if recursive:
        for root_dir, _, files in os.walk(input_path):
            for f in files:
                if f.lower().endswith(".nessus"):
                    nessus_files.append(os.path.join(root_dir, f))
    else:
        nessus_files = [
            os.path.join(input_path, f)
            for f in os.listdir(input_path)
            if f.lower().endswith(".nessus")
        ]

    if not nessus_files:
        raise FileNotFoundError(f"No .nessus files found in directory: {input_path}")
    return nessus_files


def main():
    parser = argparse.ArgumentParser(
        description="Convert Nessus (.nessus) files to CSV, including compliance fields."
    )
    parser.add_argument(
        "-i", "--input",
        required=True,
        help="Path to a .nessus file or a directory containing .nessus files."
    )
    parser.add_argument(
        "-o", "--output",
        default="nessus_results.csv",
        help="Output CSV file path (default: nessus_results.csv)"
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Recursively search for .nessus files in subdirectories (if input is a directory)."
    )

    args = parser.parse_args()

    try:
        files = collect_nessus_inputs(args.input, recursive=args.recursive)
    except Exception as e:
        print(f"[!] {e}", file=sys.stderr)
        sys.exit(1)

    fieldnames = [
        "IP Address", "Port", "Protocol", "Service",
        "Plugin ID", "Plugin Name", "Plugin Family",
        "Severity", "Risk Factor",
        "Is Compliance Item", "Compliance Result",
        "Compliance Check Name", "Compliance Check ID",
        "Compliance Policy Value", "Compliance Actual Value",
        "Compliance Solution", "Compliance Info",
        "Compliance Reference", "Compliance Benchmark Name",
        "Compliance Benchmark Version", "Compliance Benchmark Profile",
        "Compliance Full ID", "Compliance Control ID",
        "Compliance See Also", "Compliance Source",
        "Compliance Audit File", "Compliance Functional ID",
        "Compliance Informational ID",
    ]

    total_rows = 0
    with open(args.output, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for path in files:
            rows = parse_nessus_file(path)
            for r in rows:
                writer.writerow(r)
            total_rows += len(rows)

    print(f"[+] Wrote {total_rows} rows to '{args.output}' from {len(files)} file(s).")


if __name__ == "__main__":
    main()
