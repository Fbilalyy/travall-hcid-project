import graphviz

def create_hta1():
    """Method 1: Photo Gallery → Share → Mail"""
    g = graphviz.Digraph('HTA1', format='png')
    g.attr(rankdir='TB', dpi='200', bgcolor='white', pad='0.5')
    g.attr('node', shape='box', style='rounded,filled', fillcolor='#F0F4FF',
           fontname='Helvetica', fontsize='11', color='#2563EB', penwidth='1.5')
    g.attr('edge', color='#374151', penwidth='1.2')

    # Level 0
    g.node('0', '0. Share a photo via\nemail on iPhone', fillcolor='#DBEAFE', fontsize='13', penwidth='2')

    # Plan 0
    g.node('p0', 'Plan 0: Do 1 → 2 → 3 → 4 in order',
           shape='plaintext', fillcolor='white', fontsize='10', fontcolor='#6B7280')

    # Level 1
    g.node('1', '1. Locate and\nopen the photo')
    g.node('2', '2. Initiate the\nsharing process')
    g.node('3', '3. Compose\nthe email')
    g.node('4', '4. Send\nthe email')

    # Plans
    g.node('p1', 'Plan 1: Do 1.1 → 1.2 → 1.3',
           shape='plaintext', fillcolor='white', fontsize='9', fontcolor='#6B7280')
    g.node('p2', 'Plan 2: Do 2.1 → 2.2',
           shape='plaintext', fillcolor='white', fontsize='9', fontcolor='#6B7280')
    g.node('p3', 'Plan 3: Do 3.1\nIf needed, also 3.2 and/or 3.3',
           shape='plaintext', fillcolor='white', fontsize='9', fontcolor='#6B7280')

    # Level 2 - under 1
    g.node('1.1', '1.1 Unlock\niPhone', fillcolor='#F9FAFB')
    g.node('1.2', '1.2 Open\nPhotos app', fillcolor='#F9FAFB')
    g.node('1.3', '1.3 Browse &\ntap desired photo', fillcolor='#F9FAFB')

    # Level 2 - under 2
    g.node('2.1', '2.1 Tap the\nShare button', fillcolor='#F9FAFB')
    g.node('2.2', '2.2 Select "Mail"\nfrom share sheet', fillcolor='#F9FAFB')

    # Level 2 - under 3
    g.node('3.1', '3.1 Enter recipient\nemail in "To" field', fillcolor='#F9FAFB')
    g.node('3.2', '3.2 Enter subject\nline (optional)', fillcolor='#FEF3C7')
    g.node('3.3', '3.3 Write message\nbody (optional)', fillcolor='#FEF3C7')

    # Level 2 - under 4
    g.node('4.1', '4.1 Tap blue\nSend button', fillcolor='#DCFCE7')

    # Edges
    g.edge('0', 'p0', style='invis')
    g.edge('p0', '1', style='invis')
    g.edges([('0', '1'), ('0', '2'), ('0', '3'), ('0', '4')])

    g.edge('1', 'p1', style='invis')
    g.edges([('1', '1.1'), ('1', '1.2'), ('1', '1.3')])

    g.edge('2', 'p2', style='invis')
    g.edges([('2', '2.1'), ('2', '2.2')])

    g.edge('3', 'p3', style='invis')
    g.edges([('3', '3.1'), ('3', '3.2'), ('3', '3.3')])

    g.edge('4', '4.1')

    # Subgraph for alignment
    with g.subgraph() as s:
        s.attr(rank='same')
        s.node('1')
        s.node('2')
        s.node('3')
        s.node('4')

    with g.subgraph() as s:
        s.attr(rank='same')
        s.node('1.1')
        s.node('1.2')
        s.node('1.3')
        s.node('2.1')
        s.node('2.2')
        s.node('3.1')
        s.node('3.2')
        s.node('3.3')
        s.node('4.1')

    return g


def create_hta2():
    """Method 2: Mail App → Compose → Attach Photo"""
    g = graphviz.Digraph('HTA2', format='png')
    g.attr(rankdir='TB', dpi='200', bgcolor='white', pad='0.5')
    g.attr('node', shape='box', style='rounded,filled', fillcolor='#F0F4FF',
           fontname='Helvetica', fontsize='11', color='#2563EB', penwidth='1.5')
    g.attr('edge', color='#374151', penwidth='1.2')

    # Level 0
    g.node('0', '0. Share a photo via\nemail on iPhone', fillcolor='#DBEAFE', fontsize='13', penwidth='2')

    # Plan 0
    g.node('p0', 'Plan 0: Do 1 → 2 → 3 → 4 in order',
           shape='plaintext', fillcolor='white', fontsize='10', fontcolor='#6B7280')

    # Level 1
    g.node('1', '1. Open Mail\n& compose')
    g.node('2', '2. Fill in\nemail details')
    g.node('3', '3. Attach\nthe photo')
    g.node('4', '4. Send\nthe email')

    # Plans
    g.node('p1', 'Plan 1: Do 1.1 → 1.2',
           shape='plaintext', fillcolor='white', fontsize='9', fontcolor='#6B7280')
    g.node('p2', 'Plan 2: Do 2.1\nIf needed, also 2.2 and/or 2.3',
           shape='plaintext', fillcolor='white', fontsize='9', fontcolor='#6B7280')
    g.node('p3', 'Plan 3: Do 3.1 → 3.2 → 3.3\nIf photo not visible, do 3.3.1 first',
           shape='plaintext', fillcolor='white', fontsize='9', fontcolor='#6B7280')

    # Level 2 - under 1
    g.node('1.1', '1.1 Open\nMail app', fillcolor='#F9FAFB')
    g.node('1.2', '1.2 Tap\nCompose button', fillcolor='#F9FAFB')

    # Level 2 - under 2
    g.node('2.1', '2.1 Enter recipient\nemail in "To" field', fillcolor='#F9FAFB')
    g.node('2.2', '2.2 Enter subject\nline (optional)', fillcolor='#FEF3C7')
    g.node('2.3', '2.3 Write message\nbody (optional)', fillcolor='#FEF3C7')

    # Level 2 - under 3
    g.node('3.1', '3.1 Tap attachment\nicon in toolbar', fillcolor='#F9FAFB')
    g.node('3.2', '3.2 Select\n"Photo Library"', fillcolor='#F9FAFB')
    g.node('3.3', '3.3 Find & select\nthe desired photo', fillcolor='#F9FAFB')

    # Level 3 - under 3.3
    g.node('p33', 'Plan 3.3: If not in recent,\ndo 3.3.1 then 3.3.2',
           shape='plaintext', fillcolor='white', fontsize='9', fontcolor='#6B7280')
    g.node('3.3.1', '3.3.1 Browse albums\nor use search', fillcolor='#FEE2E2')
    g.node('3.3.2', '3.3.2 Tap photo\nto attach', fillcolor='#F9FAFB')

    # Level 2 - under 4
    g.node('4.1', '4.1 Review email\nwith attachment', fillcolor='#F9FAFB')
    g.node('4.2', '4.2 Tap blue\nSend button', fillcolor='#DCFCE7')

    # Edges
    g.edge('0', 'p0', style='invis')
    g.edges([('0', '1'), ('0', '2'), ('0', '3'), ('0', '4')])

    g.edge('1', 'p1', style='invis')
    g.edges([('1', '1.1'), ('1', '1.2')])

    g.edge('2', 'p2', style='invis')
    g.edges([('2', '2.1'), ('2', '2.2'), ('2', '2.3')])

    g.edge('3', 'p3', style='invis')
    g.edges([('3', '3.1'), ('3', '3.2'), ('3', '3.3')])

    g.edge('3.3', 'p33', style='invis')
    g.edges([('3.3', '3.3.1'), ('3.3', '3.3.2')])

    g.edges([('4', '4.1'), ('4', '4.2')])

    # Alignment
    with g.subgraph() as s:
        s.attr(rank='same')
        s.node('1')
        s.node('2')
        s.node('3')
        s.node('4')

    with g.subgraph() as s:
        s.attr(rank='same')
        s.node('1.1')
        s.node('1.2')
        s.node('2.1')
        s.node('2.2')
        s.node('2.3')
        s.node('3.1')
        s.node('3.2')
        s.node('3.3')
        s.node('4.1')
        s.node('4.2')

    with g.subgraph() as s:
        s.attr(rank='same')
        s.node('3.3.1')
        s.node('3.3.2')

    return g


if __name__ == '__main__':
    import os
    out_dir = '/Users/fatihbilalyilmaz/PycharmProjects/IE492/IE48L'

    hta1 = create_hta1()
    hta1.render(os.path.join(out_dir, 'hta_method1'), cleanup=True)
    print("HTA Method 1 diagram created")

    hta2 = create_hta2()
    hta2.render(os.path.join(out_dir, 'hta_method2'), cleanup=True)
    print("HTA Method 2 diagram created")
