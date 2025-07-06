import re
import json

def parse_pcf_to_json_grouped(file_path):
    with open(file_path, 'r', encoding='ISO-8859-1') as file:
        content = file.readlines()
    
    data = {
        "PIPE": [],
        "WELD": [],
        "VALVE": [],
        # Add more categories if needed
    }
    
    current_section = None
    current_item = None
    
    for line in content:
        line = line.strip()
        
        # Detect new section like PIPE or WELD
        if line.startswith("PIPE"):
            if current_item:
                data[current_section].append(current_item)
            current_section = "PIPE"
            current_item = {
                "END-POINT": [],
                "FABRICATION-ITEM": None,
                "PIPING-SPEC": None,
                "TRACING-SPEC": None,
                "COMPONENT-ATTRIBUTES": [],
                "WEIGHT": None,
                "CONTINUATION": None
            }
        
        elif line.startswith("WELD"):
            if current_item:
                data[current_section].append(current_item)
            current_section = "WELD"
            current_item = {
                "END-POINT": [],
                "FABRICATION-ITEM": None,
                "PIPING-SPEC": None,
                "TRACING-SPEC": None,
                "SKEY": None,
                "COMPONENT-ATTRIBUTES": [],
                "CONTINUATION": None
            }
        elif line.startswith("VALVE"):
            if current_item:
                data[current_section].append(current_item)
            current_section = "VALVE"
            current_item = {
                "END-POINT": [],
                "FABRICATION-ITEM": None,
                "PIPING-SPEC": None,
                "TRACING-SPEC": None,
                "SKEY": None,
                "COMPONENT-ATTRIBUTES": [],
                "CONTINUATION": None
            }
        
        # Extract END-POINT
        elif line.startswith("END-POINT") and current_item is not None:
            coords = re.findall(r"[-+]?\d*\.\d+|\d+", line)
            current_item["END-POINT"].append(tuple(map(float, coords)))
        
        # Extract FABRICATION-ITEM, PIPING-SPEC, TRACING-SPEC
        elif line.startswith("FABRICATION-ITEM") and current_item is not None:
            current_item["FABRICATION-ITEM"] = line
        elif line.startswith("PIPING-SPEC") and current_item is not None:
            current_item["PIPING-SPEC"] = line.split()[-1]
        elif line.startswith("TRACING-SPEC") and current_item is not None:
            current_item["TRACING-SPEC"] = line
        
        # Extract COMPONENT-ATTRIBUTES
        elif line.startswith("COMPONENT-ATTRIBUTE") and current_item is not None:
            current_item["COMPONENT-ATTRIBUTES"].append(line.split()[-1])
        
        # Extract SKEY
        elif line.startswith("SKEY") and current_item is not None:
            current_item["SKEY"] = line.split()[-1]
        
        # Extract weight inside pipe block
        elif line.startswith("WEIGHT") and current_item is not None:
            current_item["WEIGHT"] = float(line.split()[-1])
        
        # Extract CONTINUATION
        elif line.startswith("CONTINUATION") and current_item is not None:
            current_item["CONTINUATION"] = line
    
    # Add the last item if exists
    if current_item:
        data[current_section].append(current_item)
    
    return data

# Example usage to parse the PCF file and get the result in JSON format
file_path = r'C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\Projetos\2 - Em Verificacao\Agua de Servico - BdB200301-0120-V-MC0008\DN3-ASR-439-002-C1B.pcf'
parsed_data_json = json.dumps(parse_pcf_to_json_grouped(file_path), indent=4)

# Save the JSON output to a file
output_json_file_path = r'C:\Users\dboliveira\OneDrive - BRASS DO BRASIL\1-EXECUCAO DE PROJETO\Projetos\2 - Em Verificacao\Agua de Servico - BdB200301-0120-V-MC0008\DN3-ASR-439-002-C1B.json'
save_json_to_file(parsed_data_json, output_json_file_path)

# Return the path for user to download
output_json_file_path



"""
BOLT
CAP
COUPLING
ELBOW
END-CONNECTION-EQUIPMENT
END-CONNECTION-PIPELINE
END-POSITION-CLOSED
END-POSITION-NULL
FLANGE
FLOW-ARROW
GASKET
INSTRUMENT
ISOGEN-FILES  ISOCONFIG.FLS
ITEM-CODE  30018651
ITEM-CODE  30027454
ITEM-CODE  30027455
ITEM-CODE  30027467
ITEM-CODE  30027514
ITEM-CODE  30027515
ITEM-CODE  30027527
ITEM-CODE  30027885 
ITEM-CODE  30028609
ITEM-CODE  30028610
ITEM-CODE  30028683
ITEM-CODE  30029422
ITEM-CODE  30029431
ITEM-CODE  30029432
ITEM-CODE  30029615
ITEM-CODE  30029759
ITEM-CODE  30030002
ITEM-CODE  30030005
ITEM-CODE  30030174
ITEM-CODE  30030225
ITEM-CODE  30030279
ITEM-CODE  30030318
ITEM-CODE  30031797
ITEM-CODE  30032135
ITEM-CODE  30032137
ITEM-CODE  30032275
ITEM-CODE  30032276
ITEM-CODE  30036915
ITEM-CODE  30036918
ITEM-CODE  30036919
ITEM-CODE  30041137
ITEM-CODE  30041340
ITEM-CODE  30041862
ITEM-CODE  30041881
ITEM-CODE  30041884
ITEM-CODE  30041889
ITEM-CODE  30057324
ITEM-CODE  5fd1dc8e-7197-42f8-8bda-7d231028e19f
ITEM-CODE  7e929fca-976b-479f-b5ff-5cb5d537fe31
ITEM-CODE  7ee876e3-a462-47e9-a0a6-eaaad35ba3c7
ITEM-CODE  A1B-13803
ITEM-CODE  A1B-15788
ITEM-CODE  Buttweld-2.5
ITEM-CODE  Buttweld-3
ITEM-CODE  C1B-1239
ITEM-CODE  C1B-13803
ITEM-CODE  C1B-15788
ITEM-CODE  C1B-15854
ITEM-CODE  C1B-16045
ITEM-CODE  C1B-16579
ITEM-CODE  C1B-16647
ITEM-CODE  C1B-16931
ITEM-CODE  C1B-21197
ITEM-CODE  C1B-23207
ITEM-CODE  C1B-27003
ITEM-CODE  C1B-27005
ITEM-CODE  C1B-27266-152.4
ITEM-CODE  C1B-27268-171.45
ITEM-CODE  C1B-27268-95.25
ITEM-CODE  C1B-27281-69.85
ITEM-CODE  C1B-27283-82.55
ITEM-CODE  C1B-27284-95.25
ITEM-CODE  C1B-27304
ITEM-CODE  C1B-27322
ITEM-CODE  C1B-27323
ITEM-CODE  C1B-27325
ITEM-CODE  C1B-27351
ITEM-CODE  C1B-27353
ITEM-CODE  C1B-27425
ITEM-CODE  C1B-28003
ITEM-CODE  C1B-28006
ITEM-CODE  C1B-28063
ITEM-CODE  C1B-28066
ITEM-CODE  C1B-28069
ITEM-CODE  C1B-28381
ITEM-CODE  C1B-28391
ITEM-CODE  C1B-29842
ITEM-CODE  C1B-29844
ITEM-CODE  C1B-29858
ITEM-CODE  C1B-30903
ITEM-CODE  C1B-31437
ITEM-CODE  C1B-32905
ITEM-CODE  C1B-9306
ITEM-CODE  SlipOn-2.5
ITEM-CODE  SlipOn-3
ITEM-CODE  Tapweld-3
ITEM-CODE  Thread-0.75
ITEM-CODE  Thread-1
ITEM-CODE  Thread-1.25
ITEM-CODE  Thread-1.5
ITEM-CODE  Thread-2
ITEM-CODE  c0b54646-0fd1-4b43-8a58-032dc8fd2b1f
ITEM-CODE  c3a12be2-41fe-48a6-a02b-8bacbb0587c7
MATERIALS
MISC-COMPONENT
OLET
PIPE
PIPELINE-REFERENCE    DN3"-ASR-439-002-C1B-1
PIPELINE-REFERENCE    DN3"-ASR-439-002-C1B-2
PIPELINE-REFERENCE    DN3"-ASR-439-002-C1B-3
PIPELINE-REFERENCE    DN3"-ASR-439-002-C1B-4
PIPELINE-REFERENCE    DN3"-ASR-439-002-C1B-5
PIPELINE-REFERENCE    DN3"-ASR-439-002-C1B-6
REDUCER-CONCENTRIC
REDUCER-ECCENTRIC
TEE
UNION
UNITS-BOLT-DIA  INCH
UNITS-BOLT-LENGTH  MM
UNITS-BORE  INCH
UNITS-CO-ORDS  MM
UNITS-WEIGHT  KGS
VALVE
WELD
"""