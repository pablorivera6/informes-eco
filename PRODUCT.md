# PRODUCT.md — Informes Diarios ECO

## Product Purpose
Internal daily report automation tool for construction projects. Field engineers and project managers at Protección Catódica de Colombia (PCC) use this to generate formal daily progress reports for the client Ecopetrol, replacing manual Excel work. The tool parses FastField form submissions and writes data into a strict Ecopetrol-format Excel template.

## Register
product

## Users
- **Field project managers** (primary): semi-technical, working from laptops in field offices or trailers on oil field sites in Colombia. May be in bright outdoor light or dim site offices. Using this once per day, every workday, under moderate time pressure.
- **Data engineers / administrators** (secondary): setting up and troubleshooting the tool.

## Scene sentence
A project manager in a dusty field office trailer near Cusiana oilfield, end of shift, tired, uploading today's FastField submission and the last report, filling in quantities for the items they worked on, and hitting generate to get the Excel ready for the 6pm email to Ecopetrol.

## Brand
- **Company**: Protección Catódica de Colombia (PCC)
- **Client**: Ecopetrol (Colombia's national oil company)
- **Primary color**: #BE1E2D (PCC deep crimson red)
- **Client color**: Ecopetrol uses dark green #003B2C and yellow #F5D002
- **Tone**: Professional, precise, no-nonsense. Industrial. Not playful.
- **No emojis** anywhere in the interface.
- **Language**: Spanish

## Anti-references
- Generic SaaS dashboards with teal/blue gradients
- Startup-y light mode forms
- Heavy glassmorphism
- Overly animated UIs
- Anything that feels like a consumer app

## Strategic principles
1. **Speed over delight**: the user does this daily. Every extra click is friction. Defaults should be smart.
2. **Trust the format**: Ecopetrol's Excel is sacred. The UI's job is to fill it correctly, not to reinterpret it.
3. **Clarity in dense data**: 342 contract items across 3 project sections. The UI must make filtering and selection fast.
4. **Error prevention**: wrong data in the wrong cell has real consequences. Visual confirmation of what will change is mandatory.
