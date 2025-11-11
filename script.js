// Global variables
let treeData = null;
let root = null;
let svg = null;
let g = null;
let zoom = null;
let nodeMap = new Map();

// Password Protection
const SITE_PASSWORD = 'Technoplus2024'; // Change this to your desired password

// Check authentication status
function checkAuthentication() {
    const isAuthenticated = sessionStorage.getItem('authenticated') === 'true';
    return isAuthenticated;
}

// Show login modal
function showLoginModal() {
    const loginModal = document.getElementById('loginModal');
    const mainContent = document.getElementById('mainSiteContent');
    
    if (loginModal) {
        loginModal.classList.remove('hidden');
    }
    if (mainContent) {
        mainContent.classList.add('hidden');
    }
}

// Hide login modal and show main content
function hideLoginModal() {
    const loginModal = document.getElementById('loginModal');
    const mainContent = document.getElementById('mainSiteContent');
    
    if (loginModal) {
        loginModal.classList.add('hidden');
    }
    if (mainContent) {
        mainContent.classList.remove('hidden');
    }
}

// Handle login form submission
function handleLogin(event) {
    event.preventDefault();
    
    const passwordInput = document.getElementById('passwordInput');
    const errorMessage = document.getElementById('errorMessage');
    const enteredPassword = passwordInput.value.trim();
    
    // Clear previous error
    if (errorMessage) {
        errorMessage.classList.remove('show');
        errorMessage.textContent = '';
    }
    
    // Check password
    if (enteredPassword === SITE_PASSWORD) {
        // Set authentication
        sessionStorage.setItem('authenticated', 'true');
        
        // Hide login modal and show main content
        hideLoginModal();
        
        // Load data after authentication
        loadData();
    } else {
        // Show error message
        if (errorMessage) {
            errorMessage.textContent = 'Incorrect password. Please try again.';
            errorMessage.classList.add('show');
        }
        
        // Clear password input
        passwordInput.value = '';
        passwordInput.focus();
        
        // Shake animation
        const loginBox = document.querySelector('.login-box');
        if (loginBox) {
            loginBox.style.animation = 'shake 0.5s';
            setTimeout(() => {
                loginBox.style.animation = '';
            }, 500);
        }
    }
}

// Initialize authentication on page load
document.addEventListener('DOMContentLoaded', function() {
    // Check if user is already authenticated
    if (checkAuthentication()) {
        hideLoginModal();
        // Load data after a short delay to ensure DOM is ready
        setTimeout(() => {
            loadData();
        }, 100);
    } else {
        showLoginModal();
    }
    
    // Set up login form handler
    const loginForm = document.getElementById('loginForm');
    if (loginForm) {
        loginForm.addEventListener('submit', handleLogin);
    }
    
    // Allow Enter key to submit
    const passwordInput = document.getElementById('passwordInput');
    if (passwordInput) {
        passwordInput.addEventListener('keypress', function(e) {
            if (e.key === 'Enter') {
                handleLogin(e);
            }
        });
    }
});

// Load data from embedded JavaScript file
function loadData() {
    // Wait a bit for data.js to load if needed
    const checkData = () => {
        try {
            // Access global allData from data.js (declared as var, so it's on window)
            let data = null;
            
            // Check multiple ways to access allData
            if (typeof window !== 'undefined' && window.allData) {
                data = window.allData;
            } else if (typeof allData !== 'undefined') {
                data = allData;
                window.allData = data; // Store on window for consistency
            }
            
            if (data && Array.isArray(data) && data.length > 0) {
                console.log('Data loaded:', data.length, 'items');
                window.allData = data;
                processData();
                initializeDiagram();
                updateStats();
                populateFilters();
            } else {
                // Fallback: try to fetch from JSON file
                console.log('Trying to fetch data.json...');
                fetch('data.json')
                    .then(response => {
                        if (!response.ok) throw new Error('Failed to fetch');
                        return response.json();
                    })
                    .then(data => {
                        console.log('Data loaded from JSON:', data.length, 'items');
                        window.allData = data;
                        processData();
                        initializeDiagram();
                        updateStats();
                        populateFilters();
                    })
                    .catch(error => {
                        console.error('Error loading data:', error);
                        document.getElementById('mindmap').innerHTML = 
                            '<div style="padding: 20px; text-align: center; color: #666;">Error loading data. Please ensure data.js exists.<br><small>' + error.message + '</small></div>';
                    });
            }
        } catch (error) {
            console.error('Error loading data:', error);
            document.getElementById('mindmap').innerHTML = 
                '<div style="padding: 20px; text-align: center; color: #666;">Error loading data: ' + error.message + '</div>';
        }
    };
    
    // Check immediately, and if no data, wait a bit and check again
    checkData();
    if (!window.allData && typeof allData === 'undefined') {
        setTimeout(checkData, 100);
    }
}

// Helper function to get DB copy counts for special SMDBs
function getDBCopyCounts(smdbName) {
    const normalized = smdbName.replace(/\.01$/, '');
    
    // Special case for SMDB.TN.P2
    if (normalized === 'SMDB.TN.P2') {
        return {
            'DB.TN.LXX.1B1.01': 3,
            'DB.TN.LXX.2B1.01': 1,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    // Handle SMDB.TN.L## pattern
    const match = normalized.match(/^SMDB\.TN\.L(\d+)$/);
    
    if (!match) return null;
    
    const level = parseInt(match[1]);
    
    if (level === 1) {
        return {
            'DB.TN.LXX.1B1.01': 5,
            'DB.TN.LXX.2B1.01': 3,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    if (level === 23) {
        return {
            'DB.TN.LXX.2B1.01': 2,
            'DB.TN.LXX.1B1.01': 1,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    if ((level >= 2 && level <= 22) || (level >= 24 && level <= 47)) {
        return {
            'DB.TN.LXX.2B1.01': 4,
            'DB.TN.LXX.1B1.01': 4,
            'DB.TN.LXX.3B1.01': 1
        };
    }
    
    return null;
}

function processData() {
    const allData = window.allData || (typeof allData !== 'undefined' ? allData : []);
    if (!allData || allData.length === 0) {
        console.error('No data available');
        return;
    }
    nodeMap.clear();
    
    // Create RMU node as root
    const rmuNode = {
        name: 'RMU',
        id: 'RMU',
        kind: 'RMU',
        children: []
    };
    
    // Create MDB nodes
    const mdbList = ['MDB1', 'MDB2', 'MDB3', 'MDB4'];
    const mdbNodes = {};
    
    mdbList.forEach(mdbName => {
        // Find corresponding data item for this MDB
        const mdbData = allData.find(d => (d.Itemdrop || d.MDB) === mdbName);
        
        const mdbNode = {
            name: mdbName,
            id: mdbName,
            kind: 'MDB',
            load: mdbData ? (mdbData.Load || '0 kW') : '0 kW',
            estimate: mdbData ? (mdbData.Estimate || 0) : 0,
            mdb: mdbName,
            data: mdbData || null, // Attach data so it's clickable
            children: []
        };
        rmuNode.children.push(mdbNode);
        mdbNodes[mdbName] = mdbNode;
    });
    
    // Process all items - build complete tree structure
    const specialDBs = ['DB.TN.LXX.1B1.01', 'DB.TN.LXX.2B1.01', 'DB.TN.LXX.3B1.01'];
    const itemMap = new Map();
    const processedItems = new Set();
    
    // First pass: create all node structures
    const allDataArray = window.allData || [];
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        
        // Skip if itemName matches an MDB node (these are reference entries, not actual tree nodes)
        if (mdbList.includes(itemName)) return;
        
        // Skip if no itemName or if it's a special/reserved name
        if (!itemName || itemName === 'RMU' || specialDBs.includes(itemName)) return;
        
        const kind = item.KIND || 'Unknown';
        const load = item.Load || '0 kW';
        const estimate = item.Estimate || 0;
        const mdb = item.MDB || '';
        
        const nodeData = {
            name: itemName,
            id: itemName,
            kind: kind,
            load: load,
            estimate: estimate,
            mdb: mdb,
            data: item,
            children: []
        };
        
        itemMap.set(itemName, nodeData);
    });
    
    // Second pass: establish relationships
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        
        // Skip if itemName matches an MDB node (these are reference entries, not actual tree nodes)
        if (mdbList.includes(itemName)) return;
        
        // Skip if no itemName or if it's a special/reserved name
        if (!itemName || itemName === 'RMU' || specialDBs.includes(itemName)) return;
        if (processedItems.has(itemName)) return;
        
        const fedFrom = item['FED FROM'] || '';
        const mdb = item.MDB || '';
        const nodeData = itemMap.get(itemName);
        
        if (!nodeData) return;
        
        // Handle FED FROM relationships
        if (fedFrom && fedFrom.trim() !== '') {
            const parents = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU');
            
            if (parents.length > 0) {
                // Use first parent for tree structure
                const parent = parents[0];
                
                if (parent.startsWith('MDB') && mdbNodes[parent]) {
                    // Direct connection to MDB
                    if (!mdbNodes[parent].children.find(c => c.id === itemName)) {
                        mdbNodes[parent].children.push(nodeData);
                        processedItems.add(itemName);
                    }
                } else {
                    // Connect to parent item
                    let parentNode = itemMap.get(parent);
                    if (!parentNode) {
                        // Create placeholder parent
                        parentNode = {
                            name: parent,
                            id: parent,
                            kind: parent.startsWith('BB') ? 'BUS BAR RAISER' : 
                                  parent.startsWith('SMDB') ? 'SMDB' : 
                                  parent.startsWith('DB') ? 'DB' : 'Parent',
                            children: []
                        };
                        itemMap.set(parent, parentNode);
                        
                        // Try to connect parent to its MDB
                        const parentItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === parent);
                        if (parentItem) {
                            const parentMDB = parentItem.MDB || '';
                            const parentFedFrom = parentItem['FED FROM'] || '';
                            if (parentMDB && mdbNodes[parentMDB]) {
                                if (!mdbNodes[parentMDB].children.find(c => c.id === parent)) {
                                    mdbNodes[parentMDB].children.push(parentNode);
                                }
                            } else if (parentFedFrom.includes('MDB')) {
                                const mdbMatch = parentFedFrom.match(/MDB\d/);
                                if (mdbMatch && mdbNodes[mdbMatch[0]]) {
                                    if (!mdbNodes[mdbMatch[0]].children.find(c => c.id === parent)) {
                                        mdbNodes[mdbMatch[0]].children.push(parentNode);
                                    }
                                }
                            }
                        }
                    }
                    
                    if (!parentNode.children.find(c => c.id === itemName)) {
                        parentNode.children.push(nodeData);
                        processedItems.add(itemName);
                    }
                }
            } else if (mdb && mdbNodes[mdb]) {
                // No parent specified, connect to MDB
                if (!mdbNodes[mdb].children.find(c => c.id === itemName)) {
                    mdbNodes[mdb].children.push(nodeData);
                    processedItems.add(itemName);
                }
            }
        } else if (mdb && mdbNodes[mdb]) {
            // No FED FROM, connect to MDB
            if (!mdbNodes[mdb].children.find(c => c.id === itemName)) {
                mdbNodes[mdb].children.push(nodeData);
                processedItems.add(itemName);
            }
        }
    });
    
    // Handle special DBs with grouping
    // First, process special DBs from their FED FROM fields
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        if (!specialDBs.includes(itemName)) return;
        
        const fedFrom = item['FED FROM'] || '';
        const kind = item.KIND || 'Unknown';
        const load = item.Load || '0 kW';
        const estimate = item.Estimate || 0;
        const mdb = item.MDB || '';
        
        const smdbList = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU');
        
        smdbList.forEach(smdbName => {
            const copyCounts = getDBCopyCounts(smdbName);
            if (copyCounts && copyCounts[itemName]) {
                const count = copyCounts[itemName];
                
                // Find or create SMDB node
                let smdbNode = itemMap.get(smdbName);
                if (!smdbNode) {
                    smdbNode = {
                        name: smdbName,
                        id: smdbName,
                        kind: 'SMDB',
                        children: []
                    };
                    itemMap.set(smdbName, smdbNode);
                    
                    // Connect SMDB to its parent (find from data)
                    const smdbItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === smdbName);
                    if (smdbItem) {
                        const smdbFedFrom = smdbItem['FED FROM'] || '';
                        if (smdbFedFrom.includes('MDB')) {
                            const mdbMatch = smdbFedFrom.match(/MDB\d/);
                            if (mdbMatch && mdbNodes[mdbMatch[0]]) {
                                if (!mdbNodes[mdbMatch[0]].children.find(c => c.id === smdbName)) {
                                    mdbNodes[mdbMatch[0]].children.push(smdbNode);
                                }
                            }
                        }
                    }
                }
                
                // Create group node
                const groupNode = {
                    name: `${itemName} (${count} copies)`,
                    id: `${smdbName}_${itemName}_group`,
                    kind: kind,
                    load: load,
                    estimate: estimate,
                    mdb: mdb,
                    data: { ...item, copyCount: count, fedFromSMDB: smdbName, isGroup: true },
                    groupCount: count,
                    collapsed: true,
                    children: []
                };
                
                smdbNode.children.push(groupNode);
            }
        });
    });
    
    // Second, process SMDBs that have special copy counts but may not be in FED FROM
    // This handles cases like SMDB.TN.P2.01
    allDataArray.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        if (!itemName.startsWith('SMDB.TN.')) return;
        
        const copyCounts = getDBCopyCounts(itemName);
        if (!copyCounts) return;
        
        // Find or create SMDB node
        let smdbNode = itemMap.get(itemName);
        if (!smdbNode) {
            smdbNode = {
                name: itemName,
                id: itemName,
                kind: 'SMDB',
                children: []
            };
            itemMap.set(itemName, smdbNode);
            
            // Connect SMDB to its parent (find from data)
            const smdbFedFrom = item['FED FROM'] || '';
            if (smdbFedFrom.includes('MDB')) {
                const mdbMatch = smdbFedFrom.match(/MDB\d/);
                if (mdbMatch && mdbNodes[mdbMatch[0]]) {
                    if (!mdbNodes[mdbMatch[0]].children.find(c => c.id === itemName)) {
                        mdbNodes[mdbMatch[0]].children.push(smdbNode);
                    }
                }
            }
        }
        
        // Create group nodes for each special DB type
        specialDBs.forEach(dbType => {
            if (copyCounts[dbType]) {
                const count = copyCounts[dbType];
                
                // Check if group node already exists
                const groupId = `${itemName}_${dbType}_group`;
                if (smdbNode.children.find(c => c.id === groupId)) return;
                
                // Find the DB item data
                const dbItem = allDataArray.find(d => (d.Itemdrop || d.MDB) === dbType);
                const kind = dbItem ? (dbItem.KIND || 'Unknown') : 'DB';
                const load = dbItem ? (dbItem.Load || '0 kW') : '0 kW';
                const estimate = dbItem ? (dbItem.Estimate || 0) : 0;
                const mdb = dbItem ? (dbItem.MDB || '') : '';
                
                // Create group node
                const groupNode = {
                    name: `${dbType} (${count} copies)`,
                    id: groupId,
                    kind: kind,
                    load: load,
                    estimate: estimate,
                    mdb: mdb,
                    data: dbItem ? { ...dbItem, copyCount: count, fedFromSMDB: itemName, isGroup: true } : { copyCount: count, fedFromSMDB: itemName, isGroup: true },
                    groupCount: count,
                    collapsed: true,
                    children: []
                };
                
                smdbNode.children.push(groupNode);
            }
        });
    });
    
    treeData = rmuNode;
    root = d3.hierarchy(treeData);
    root.x0 = 0;
    root.y0 = 0;
}

function initializeDiagram() {
    const container = d3.select('#mindmap');
    container.selectAll('*').remove();
    
    const width = container.node().offsetWidth || 1200;
    const height = Math.max(600, window.innerHeight * 0.6);
    
    // Create SVG
    svg = container
        .append('svg')
        .attr('width', width)
        .attr('height', height)
        .attr('viewBox', `0 0 ${width} ${height}`);
    
    // Create zoom behavior
    zoom = d3.zoom()
        .scaleExtent([0.1, 3])
        .on('zoom', (event) => {
            g.attr('transform', event.transform);
        });
    
    svg.call(zoom);
    
    // Create main group
    g = svg.append('g');
    
    // Create tree layout
    const tree = d3.tree()
        .size([height - 100, width - 200])
        .separation((a, b) => (a.parent === b.parent ? 1 : 1.5) / a.depth);
    
    // Update function
    function update(source) {
        const treeData = tree(root);
        const nodes = treeData.descendants();
        const links = treeData.links();
        
        // Normalize for fixed-depth
        nodes.forEach(d => {
            d.y = d.depth * 180;
        });
        
        // Update nodes
        const node = g.selectAll('g.node')
            .data(nodes, d => d.data.id);
        
        const nodeEnter = node.enter()
            .append('g')
            .attr('class', 'node')
            .attr('transform', d => `translate(${source.y0},${source.x0})`)
            .on('click', click);
        
        // Add circles for nodes
        nodeEnter.append('circle')
            .attr('r', d => getNodeRadius(d))
            .attr('fill', d => getNodeColor(d.data.kind))
            .attr('stroke', d => getNodeBorderColor(d.data.kind))
            .attr('stroke-width', d => d.data.id === 'RMU' ? 3 : 2)
            .style('cursor', 'pointer');
        
        // Add labels (clickable to open detailed page)
        const labels = nodeEnter.append('text')
            .attr('dy', d => d.data.id === 'RMU' ? 5 : 4)
            .attr('x', d => {
                // Group nodes need more space for longer labels
                if (d.data.isGroup || (d.data.name && d.data.name.includes('copies'))) {
                    return d.children || d._children ? -15 : 15;
                }
                return d.children || d._children ? -13 : 13;
            })
            .attr('text-anchor', d => d.children || d._children ? 'end' : 'start')
            .attr('font-size', d => getNodeFontSize(d))
            .attr('font-weight', d => {
                if (d.data.id === 'RMU' || d.data.kind === 'MDB') return 'bold';
                // Make group nodes (with copy counts) slightly bolder for readability
                if (d.data.isGroup || (d.data.name && d.data.name.includes('copies'))) return '600';
                return 'normal';
            })
            .text(d => {
                const name = d.data.name;
                // Don't truncate group nodes (they contain copy counts)
                if (d.data.isGroup || name.includes('copies')) {
                    return name;
                }
                // Truncate other long names
                return name.length > 20 ? name.substring(0, 18) + '...' : name;
            })
            .style('pointer-events', 'all')
            .style('fill', '#333')
            .style('cursor', d => d.data.data ? 'pointer' : 'default')
            .on('click', function(event, d) {
                event.stopPropagation();
                if (d.data.data) {
                    openDetailedPage(d.data.data, d.data.name);
                }
            });
        
        // Add expand/collapse indicators
        nodeEnter.filter(d => d.children || d._children)
            .append('text')
            .attr('dy', 5)
            .attr('x', d => d.children ? -8 : -8)
            .attr('text-anchor', 'middle')
            .attr('font-size', '12px')
            .text(d => d.children ? '▼' : '▶')
            .style('fill', '#666')
            .style('pointer-events', 'none');
        
        // Update node positions
        const nodeUpdate = nodeEnter.merge(node);
        nodeUpdate.transition()
            .duration(300)
            .attr('transform', d => `translate(${d.y},${d.x})`);
        
        // Remove exiting nodes
        const nodeExit = node.exit()
            .transition()
            .duration(300)
            .attr('transform', d => `translate(${source.y},${source.x})`)
            .remove();
        
        // Update links
        const link = g.selectAll('path.link')
            .data(links, d => d.target.data.id);
        
        const linkEnter = link.enter()
            .insert('path', 'g')
            .attr('class', 'link')
            .attr('d', d => {
                const o = { x: source.x0, y: source.y0 };
                return diagonal({ source: o, target: o });
            })
            .attr('fill', 'none')
            .attr('stroke', '#999')
            .attr('stroke-width', 2);
        
        const linkUpdate = linkEnter.merge(link);
        linkUpdate.transition()
            .duration(300)
            .attr('d', diagonal);
        
        link.exit()
            .transition()
            .duration(300)
            .attr('d', d => {
                const o = { x: source.x, y: source.y };
                return diagonal({ source: o, target: o });
            })
            .remove();
        
        // Store positions for transitions
        nodes.forEach(d => {
            d.x0 = d.x;
            d.y0 = d.y;
        });
    }
    
    function diagonal(d) {
        return `M ${d.source.y},${d.source.x}
                C ${(d.source.y + d.target.y) / 2},${d.source.x}
                  ${(d.source.y + d.target.y) / 2},${d.target.x}
                  ${d.target.y},${d.target.x}`;
    }
    
    function click(event, d) {
        // Don't handle clicks on text labels (they have their own handler)
        if (event.target.tagName === 'text') {
            return;
        }
        
        // Regular click for expand/collapse
        if (d.children) {
            d._children = d.children;
            d.children = null;
        } else {
            d.children = d._children;
            d._children = null;
        }
        
        // Handle group nodes
        if (d.data.isGroup && d.data.groupCount) {
            if (d.children) {
                // Expand group
                expandGroupNode(d);
            } else {
                // Collapse group
                collapseGroupNode(d);
            }
        }
        
        // Double click on node circle opens detailed page (fallback)
        if (event.detail === 2 && d.data.data) {
            openDetailedPage(d.data.data, d.data.name);
        }
        
        update(d);
    }
    
    function expandGroupNode(d) {
        if (!d.data.groupCount || d.children) return;
        
        d.children = [];
        for (let i = 1; i <= d.data.groupCount; i++) {
            d.children.push({
                name: `Copy #${i}`,
                id: `${d.data.id}_copy${i}`,
                kind: d.data.kind,
                load: d.data.load,
                estimate: d.data.estimate,
                mdb: d.data.mdb,
                data: { ...d.data.data, copyNumber: i },
                children: null
            });
        }
    }
    
    function collapseGroupNode(d) {
        if (d.children) {
            d._children = d.children;
            d.children = null;
        }
    }
    
    // Collapse all nodes initially
    root.children.forEach(collapse);
    root.children.forEach(collapse);
    root.children.forEach(collapse);
    
    update(root);
    
    // Center the diagram
    const bounds = g.node().getBBox();
    const fullWidth = width;
    const fullHeight = height;
    const widthScale = (fullWidth - 40) / bounds.width;
    const heightScale = (fullHeight - 40) / bounds.height;
    const scale = Math.min(widthScale, heightScale, 1);
    
    const translate = [
        (fullWidth - bounds.width * scale) / 2 - bounds.x * scale,
        (fullHeight - bounds.height * scale) / 2 - bounds.y * scale
    ];
    
    g.attr('transform', `translate(${translate}) scale(${scale})`);
    
    function collapse(d) {
        if (d.children) {
            d._children = d.children;
            d._children.forEach(collapse);
            d.children = null;
        }
    }
}

function getNodeRadius(d) {
    if (d.data.id === 'RMU') return 25;
    if (d.data.kind === 'MDB' || d.data.id === 'RMU') return 20;
    if (d.data.kind === 'BUS BAR RAISER' || d.data.kind === 'SMDB') return 15;
    return 12;
}

function getNodeColor(kind) {
    const colorMap = {
        'RMU': '#667eea',
        'MDB': '#FF6B6B',
        'BUS BAR RAISER': '#4ECDC4',
        'SMDB': '#95E1D3',
        'DB': '#A8E6CF',
        'EDB': '#FFD93D',
        'EMDB': '#C7CEEA',
        'EMCC': '#FFB6C1',
        'ESMDB': '#DDA0DD',
        'MCC': '#98D8C8',
        'POWR FACTOR CORRECTOR': '#FFE66D',
        'EFP': '#F7DC6F',
        'CP': '#AED6F1',
        'EVCS': '#F8C471',
        'CHWP': '#85C1E2',
        'AHU': '#F1948A',
        'CU': '#F9E79F',
        'GSM': '#BB8FCE',
        'CBS': '#85C1E9',
        'UDB': '#F7DC6F',
        'SPF': '#A3E4D7',
        'JEF': '#FAD7A0'
    };
    return colorMap[kind] || '#D3D3D3';
}

function getNodeBorderColor(kind) {
    const colorMap = {
        'RMU': '#7D3C98',
        'MDB': '#C92A2A',
        'BUS BAR RAISER': '#0E7C7B',
        'SMDB': '#2D8659',
        'DB': '#2D8659'
    };
    return colorMap[kind] || '#808080';
}

function getNodeFontSize(d) {
    if (d.data.id === 'RMU') return '14px';
    // Group nodes (with copy counts) should be more readable
    if (d.data.isGroup || (d.data.name && d.data.name.includes('copies'))) return '11px';
    if (d.data.kind === 'MDB' || d.data.id === 'RMU') return '12px';
    if (d.data.kind === 'BUS BAR RAISER' || d.data.kind === 'SMDB') return '11px';
    return '10px';
}

function updateStats() {
    const data = window.allData || [];
    if (!data || data.length === 0) return;
    
    const totalItems = data.length;
    let totalLoad = 0;
    let totalEstimate = 0;
    
    // Find the NET TOTAL or TOTAL row from the totallist table - use that as the source of truth
    // The estimate column in totallist table should be used directly
    let tableTotalEstimate = null;
    let foundTotalRow = null;
    
    // Search through all data to find NET TOTAL or TOTAL row
    // Check both Itemdrop and MDB fields, and also check KIND field
    for (let i = data.length - 1; i >= 0; i--) {
        const item = data[i];
        const itemdrop = (item.Itemdrop || '').toString().toLowerCase().trim();
        const mdb = (item.MDB || '').toString().toLowerCase().trim();
        const kind = (item.KIND || '').toString().toLowerCase().trim();
        
        // Check if this is a NET TOTAL row
        if (itemdrop === 'net total' || mdb === 'net total' || kind === 'net total') {
            tableTotalEstimate = parseFloat(item.Estimate) || 0;
            foundTotalRow = item;
            console.log('Found NET TOTAL row:', {
                Itemdrop: item.Itemdrop,
                MDB: item.MDB,
                KIND: item.KIND,
                Estimate: tableTotalEstimate
            });
            break;
        }
    }
    
    // If NET TOTAL not found, look for TOTAL (but not "SUM AFTER TOTAL" or similar)
    if (tableTotalEstimate === null || tableTotalEstimate === 0) {
        for (let i = data.length - 1; i >= 0; i--) {
            const item = data[i];
            const itemdrop = (item.Itemdrop || '').toString().toLowerCase().trim();
            const mdb = (item.MDB || '').toString().toLowerCase().trim();
            const kind = (item.KIND || '').toString().toLowerCase().trim();
            
            // Check for exact "total" match (not "sum after total" or similar)
            if ((itemdrop === 'total' || mdb === 'total' || kind === 'total') &&
                !itemdrop.includes('sum after') && 
                !itemdrop.includes('net') &&
                !mdb.includes('sum after') &&
                !mdb.includes('net')) {
                tableTotalEstimate = parseFloat(item.Estimate) || 0;
                foundTotalRow = item;
                console.log('Found TOTAL row:', {
                    Itemdrop: item.Itemdrop,
                    MDB: item.MDB,
                    KIND: item.KIND,
                    Estimate: tableTotalEstimate
                });
                break;
            }
        }
    }
    
    // If still not found, check for rows with empty Itemdrop/MDB but high estimate (likely the total row)
    if (tableTotalEstimate === null || tableTotalEstimate === 0) {
        // Look for rows with empty Itemdrop/MDB/KIND but with a high estimate value
        // This is often how total rows are stored in Excel exports
        for (let i = data.length - 1; i >= 0; i--) {
            const item = data[i];
            const itemdrop = (item.Itemdrop || '').toString().trim();
            const mdb = (item.MDB || '').toString().trim();
            const kind = (item.KIND || '').toString().trim();
            const estimate = parseFloat(item.Estimate) || 0;
            
            // If all name fields are empty but estimate is very high (likely the total)
            if ((!itemdrop || itemdrop === '') && 
                (!mdb || mdb === '') && 
                (!kind || kind === '') &&
                estimate > 1000000) {
                tableTotalEstimate = estimate;
                foundTotalRow = item;
                console.log('Found total row (empty fields, high estimate):', {
                    Itemdrop: item.Itemdrop,
                    MDB: item.MDB,
                    KIND: item.KIND,
                    Estimate: tableTotalEstimate
                });
                break;
            }
        }
    }
    
    // Debug: Log all rows with "total" in their name to help identify the correct row
    if (!foundTotalRow) {
        console.log('Searching for total rows...');
        const totalRows = data.filter(item => {
            const itemdrop = (item.Itemdrop || '').toString().toLowerCase();
            const mdb = (item.MDB || '').toString().toLowerCase();
            return itemdrop.includes('total') || mdb.includes('total');
        });
        console.log('Rows containing "total":', totalRows.map(item => ({
            Itemdrop: item.Itemdrop,
            MDB: item.MDB,
            KIND: item.KIND,
            Estimate: item.Estimate
        })));
        
        // Also check last few rows with high estimates
        console.log('Last 5 rows:', data.slice(-5).map(item => ({
            Itemdrop: item.Itemdrop,
            MDB: item.MDB,
            KIND: item.KIND,
            Estimate: item.Estimate
        })));
    }
    
    // Calculate total load (sum individual items for load)
    data.forEach(item => {
        const itemName = item.Itemdrop || item.MDB || '';
        
        // Skip items with no name
        if (!itemName || itemName.trim() === '') {
            return;
        }
        
        // Skip MDB reference entries
        if (['MDB1', 'MDB2', 'MDB3', 'MDB4'].includes(itemName)) {
            return;
        }
        
        const loadStr = item.Load || '0 kW';
        const loadValue = parseFloat(loadStr.replace(/[^\d.]/g, '')) || 0;
        totalLoad += loadValue;
    });
    
    // Use table total estimate directly from totallist table (includes special DB values)
    if (tableTotalEstimate !== null && tableTotalEstimate > 0) {
        totalEstimate = tableTotalEstimate;
        console.log('Using table total estimate from totallist:', totalEstimate);
    } else {
        // Fallback: if no total row found, calculate from all items
        console.warn('NET TOTAL or TOTAL row not found in data, calculating from items');
        data.forEach(item => {
            const itemName = item.Itemdrop || item.MDB || '';
            if (!itemName || itemName.trim() === '') return;
            if (['MDB1', 'MDB2', 'MDB3', 'MDB4'].includes(itemName)) return;
            totalEstimate += parseFloat(item.Estimate) || 0;
        });
    }

    document.getElementById('totalItems').textContent = totalItems;
    document.getElementById('totalLoad').textContent = totalLoad.toFixed(2) + ' kW';
    document.getElementById('totalEstimate').textContent = formatNumber(totalEstimate);
}

function populateFilters() {
    const kindFilter = document.getElementById('kindFilter');
    const data = window.allData || [];
    if (!data || data.length === 0) return;
    const kinds = [...new Set(data.map(item => item.KIND).filter(Boolean))].sort();
    
    kinds.forEach(kind => {
        const option = document.createElement('option');
        option.value = kind;
        option.textContent = kind;
        kindFilter.appendChild(option);
    });
}

function formatNumber(num) {
    if (!num) return 'N/A';
    return new Intl.NumberFormat('en-US', {
        minimumFractionDigits: 2,
        maximumFractionDigits: 2
    }).format(num);
}

function showDetails(item) {
    const panel = document.getElementById('details-panel');
    const content = document.getElementById('detailsContent');
    
    const isCopy = item.copyNumber && item.fedFromSMDB;
    const itemName = item.Itemdrop || item.MDB || 'N/A';
    const fedFrom = isCopy ? item.fedFromSMDB : (item['FED FROM'] || 'N/A');
    
    content.innerHTML = `
        <div class="detail-item">
            <div class="detail-label">Item Name:</div>
            <div class="detail-value">${itemName}${isCopy ? ` (Copy #${item.copyNumber})` : ''}</div>
        </div>
        ${isCopy ? `
        <div class="detail-item">
            <div class="detail-label">Copy Number:</div>
            <div class="detail-value">${item.copyNumber}</div>
        </div>
        ` : ''}
        <div class="detail-item">
            <div class="detail-label">Kind:</div>
            <div class="detail-value">${item.KIND || 'N/A'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">MDB:</div>
            <div class="detail-value">${item.MDB || 'N/A'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Fed From:</div>
            <div class="detail-value">${fedFrom}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Load:</div>
            <div class="detail-value">${item.Load || 'N/A'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Number of Items:</div>
            <div class="detail-value">${item['NO OF ITEMS'] || 'N/A'}</div>
        </div>
        <div class="detail-item">
            <div class="detail-label">Estimate:</div>
            <div class="detail-value">${formatNumber(item.Estimate)}</div>
        </div>
    `;
    
    panel.classList.remove('hidden');
}

// Open detailed page with metadata and Excel sheet data
function openDetailedPage(item, nodeName, sourceView = 'main') {
    const detailedPage = document.getElementById('detailed-page');
    const detailedTitle = document.getElementById('detailed-title');
    const metadataContent = document.getElementById('metadata-content');
    const detailedDataContent = document.getElementById('detailed-data-content');
    
    // Store the source view so we know where to return
    detailedPage.setAttribute('data-source-view', sourceView);
    
    // Hide main content or MDB view based on source
    if (sourceView === 'mdb') {
        document.getElementById('mdb-view-page').classList.add('hidden');
    } else {
        document.querySelector('main > .controls').style.display = 'none';
        document.querySelector('main > .legend-panel').style.display = 'none';
        document.querySelector('main > .info-panel').style.display = 'none';
        document.getElementById('main-content').style.display = 'none';
    }
    document.getElementById('details-panel').classList.add('hidden');
    
    // Show detailed page
    detailedPage.classList.remove('hidden');
    
    // Set title
    const itemName = item ? (item.Itemdrop || item.MDB || nodeName) : nodeName || 'N/A';
    detailedTitle.textContent = itemName;
    
    // Display metadata (use item if available, otherwise create minimal metadata)
    if (item) {
        displayMetadata(item, metadataContent);
    } else {
        // Create minimal metadata for nodes without data
        metadataContent.innerHTML = `
            <table class="metadata-table">
                <tr>
                    <td>Item Name</td>
                    <td>${itemName}</td>
                </tr>
            </table>
        `;
    }
    
    // Load and display detailed data from Excel sheet
    detailedDataContent.innerHTML = '<div class="loading">Loading detailed data...</div>';
    
    // Check if this is a special DB that needs recalculation
    const specialDBs = ['DB.TN.LXX.1B1.01', 'DB.TN.LXX.2B1.01', 'DB.TN.LXX.3B1.01'];
    const isSpecialDB = specialDBs.includes(itemName);
    
    if (isSpecialDB && item) {
        // Find the parent SMDB from the tree structure
        const smdbName = findParentSMDB(itemName, item);
        if (smdbName) {
            const copyCounts = getDBCopyCounts(smdbName);
            const noOfUnits = copyCounts && copyCounts[itemName] ? copyCounts[itemName] : 1;
            // Load with recalculation
            loadDetailedDataWithRecalculation(itemName, detailedDataContent, noOfUnits);
        } else {
            // Fallback to normal loading
            loadDetailedData(itemName, detailedDataContent);
        }
    } else {
        // Normal loading (SMDBs and other items show as-is)
        loadDetailedData(itemName, detailedDataContent);
    }
}

// Find parent SMDB for a special DB
function findParentSMDB(dbName, item) {
    // Check if item has fedFromSMDB (for copy nodes)
    if (item.fedFromSMDB) {
        return item.fedFromSMDB;
    }
    
    // Check FED FROM field
    const fedFrom = item['FED FROM'] || '';
    if (fedFrom) {
        // FED FROM can contain multiple SMDBs separated by newlines
        const smdbList = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU' && !p.startsWith('MDB'));
        if (smdbList.length > 0) {
            // Return the first SMDB (or we could check which one has this DB)
            return smdbList[0];
        }
    }
    
    // Search the tree structure to find parent SMDB
    if (treeData && treeData.children) {
        for (const mdb of treeData.children) {
            if (mdb.children) {
                for (const child of mdb.children) {
                    if (child.children) {
                        for (const grandchild of child.children) {
                            // Check if this is a group node containing our DB
                            if (grandchild.data && grandchild.data.isGroup && grandchild.data.name && grandchild.data.name.includes(dbName)) {
                                return child.name; // Return the SMDB name
                            }
                            // Check if this is the DB node itself
                            if (grandchild.name === dbName || grandchild.id === dbName) {
                                return child.name; // Return the SMDB name
                            }
                        }
                    }
                }
            }
        }
    }
    
    return null;
}

// Display metadata in table format
function displayMetadata(item, container) {
    const isCopy = item.copyNumber && item.fedFromSMDB;
    const itemName = item.Itemdrop || item.MDB || 'N/A';
    const fedFrom = isCopy ? item.fedFromSMDB : (item['FED FROM'] || 'N/A');
    
    container.innerHTML = `
        <table class="metadata-table">
            <tr>
                <td>Item Name</td>
                <td>${itemName}${isCopy ? ` (Copy #${item.copyNumber})` : ''}</td>
            </tr>
            ${isCopy ? `
            <tr>
                <td>Copy Number</td>
                <td>${item.copyNumber}</td>
            </tr>
            ` : ''}
            <tr>
                <td>Kind</td>
                <td>${item.KIND || 'N/A'}</td>
            </tr>
            <tr>
                <td>MDB</td>
                <td>${item.MDB || 'N/A'}</td>
            </tr>
            <tr>
                <td>Fed From</td>
                <td>${fedFrom}</td>
            </tr>
            <tr>
                <td>Load</td>
                <td>${item.Load || 'N/A'}</td>
            </tr>
            <tr>
                <td>Number of Items</td>
                <td>${item['NO OF ITEMS'] || 'N/A'}</td>
            </tr>
            <tr>
                <td>Estimate</td>
                <td>${formatNumber(item.Estimate)}</td>
            </tr>
        </table>
    `;
}

// Load detailed data from Excel sheet
function loadDetailedData(itemName, container) {
    // Try to fetch the Excel file
    fetch('e2.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Excel file not found');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Normal loading - show sheet as-is (no recalculation)
            loadNormalSheet(itemName, container, workbook);
        })
        .catch(error => {
            console.error('Error loading Excel file:', error);
            container.innerHTML = `
                <div class="error-message">
                    Error loading detailed data: ${error.message}
                    <br><small>Make sure e2.xlsx is in the same directory</small>
                </div>
            `;
        });
}

// Load detailed data with recalculation for special DBs
function loadDetailedDataWithRecalculation(dbName, container, noOfUnits) {
    // Try to fetch the Excel file
    fetch('e2.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error('Excel file not found');
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const sheetNames = workbook.SheetNames;
            
            // Find DB sheet
            let dbSheetName = sheetNames.find(name => 
                name === dbName || 
                name.toLowerCase() === dbName.toLowerCase()
            );
            
            if (!dbSheetName) {
                dbSheetName = sheetNames.find(name => 
                    name.toLowerCase().includes(dbName.toLowerCase()) ||
                    dbName.toLowerCase().includes(name.toLowerCase())
                );
            }
            
            if (!dbSheetName) {
                container.innerHTML = `
                    <div class="error-message">
                        No matching sheet found for "${dbName}". Available sheets: ${sheetNames.join(', ')}
                    </div>
                `;
                return;
            }
            
            const dbWorksheet = workbook.Sheets[dbSheetName];
            const dbData = XLSX.utils.sheet_to_json(dbWorksheet, { header: 1 });
            
            // Calculate estimate for this DB
            const calculation = calculateDBEstimate(dbData, noOfUnits);
            
            if (!calculation) {
                // Fallback to normal display if calculation fails
                displayDetailedTable(dbData, container, dbSheetName);
                return;
            }
            
            // Display with recalculated values
            displayDetailedTableWithRecalculation(dbData, container, dbSheetName, calculation);
        })
        .catch(error => {
            console.error('Error loading Excel file:', error);
            container.innerHTML = `
                <div class="error-message">
                    Error loading detailed data: ${error.message}
                    <br><small>Make sure e2.xlsx is in the same directory</small>
                </div>
            `;
        });
}

// Load normal sheet (non-SMDB or SMDB without special DBs)
function loadNormalSheet(itemName, container, workbook) {
    const sheetNames = workbook.SheetNames;
    let sheetName = null;
    
    // Try exact match first
    if (sheetNames.includes(itemName)) {
        sheetName = itemName;
    } else {
        // Try to find sheet that contains the item name
        sheetName = sheetNames.find(name => 
            name.toLowerCase().includes(itemName.toLowerCase()) ||
            itemName.toLowerCase().includes(name.toLowerCase())
        );
    }
    
    // If no match, try to find a sheet with similar pattern
    if (!sheetName) {
        const baseName = itemName.split('.')[0] || itemName;
        sheetName = sheetNames.find(name => 
            name.toLowerCase().includes(baseName.toLowerCase())
        );
    }
    
    // Default to first sheet if no match
    if (!sheetName && sheetNames.length > 0) {
        sheetName = sheetNames[0];
    }
    
    if (sheetName) {
        const worksheet = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
        displayDetailedTable(jsonData, container, sheetName);
    } else {
        container.innerHTML = `
            <div class="error-message">
                No matching sheet found for "${itemName}". Available sheets: ${sheetNames.join(', ')}
            </div>
        `;
    }
}


// Calculate DB estimate based on formula
function calculateDBEstimate(dbData, noOfUnits) {
    if (!dbData || dbData.length === 0) return null;
    
    const headers = dbData[0] || [];
    const rows = dbData.slice(1);
    
    // Find column indices - use same logic as display function
    const itemColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('item') || 
              h.toString().toLowerCase().includes('description'))
    );
    
    // Find AMOUNT column specifically (prioritize exact match, then partial)
    let amountColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().trim() === 'amount'
    );
    
    // If exact match not found, try partial match
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('amount') && 
                  !h.toString().toLowerCase().includes('price'))
        );
    }
    
    // Fallback to cost/estimate if amount not found
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('cost') ||
                  h.toString().toLowerCase().includes('estimate'))
        );
    }
    
    console.log('calculateDBEstimate - Column indices:', { itemColumnIndex, amountColumnIndex });
    console.log('calculateDBEstimate - Headers:', headers);
    
    if (itemColumnIndex === -1 || amountColumnIndex === -1) {
        console.warn('Could not find required columns in DB sheet');
        return null;
    }
    
    let total = 0;
    let labour = 0;
    let sumAfterLabour = 0;
    
    // Find TOTAL and LABOUR rows - check all columns for the label, then use amountColumnIndex for value
    rows.forEach((row, idx) => {
        // Check all columns for TOTAL/LABOUR labels
        let foundLabel = '';
        let labelColumnIndex = -1;
        
        for (let colIdx = 0; colIdx < row.length; colIdx++) {
            const cellValue = row[colIdx];
            if (!cellValue) continue;
            
            const cellStr = cellValue.toString().toLowerCase().trim();
            
            // Check for TOTAL row - exact match, not "NET TOTAL" or "SUM AFTER TOTAL"
            if (cellStr === 'total' && !cellStr.includes('sum after') && !cellStr.includes('net')) {
                foundLabel = 'TOTAL';
                labelColumnIndex = colIdx;
                break;
            } 
            // Check for LABOUR row
            else if (cellStr === 'labour' || cellStr === 'labor') {
                foundLabel = 'LABOUR';
                labelColumnIndex = colIdx;
                break;
            } 
            // Check for SUM AFTER LABOUR
            else if (cellStr.includes('sum after labour') || cellStr.includes('sum after labor')) {
                foundLabel = 'SUM AFTER LABOUR';
                labelColumnIndex = colIdx;
                break;
            }
        }
        
        // If we found a label, get the amount from the amount column
        if (foundLabel && amountColumnIndex !== -1 && row[amountColumnIndex] !== undefined) {
            const amountValue = parseFloat(row[amountColumnIndex]) || 0;
            
            if (foundLabel === 'TOTAL') {
                total = amountValue;
                console.log(`Found TOTAL at row ${idx}:`, row[labelColumnIndex], 'Amount column value =', amountValue, 'Full row:', row);
            } else if (foundLabel === 'LABOUR') {
                labour = amountValue;
                console.log(`Found LABOUR at row ${idx}:`, row[labelColumnIndex], 'Amount column value =', amountValue, 'Full row:', row);
            } else if (foundLabel === 'SUM AFTER LABOUR') {
                sumAfterLabour = amountValue;
                console.log(`Found SUM AFTER LABOUR at row ${idx}:`, row[labelColumnIndex], 'Amount column value =', amountValue, 'Full row:', row);
            }
        }
    });
    
    console.log('Extracted values:', { total, labour, sumAfterLabour, noOfUnits });
    
    // Validate that we have required values
    if (isNaN(total) || total === 0) {
        console.error('TOTAL is missing or zero after search. Sample rows:', rows.slice(140, 160).map((r, i) => ({
            index: 140 + i,
            item: r[itemColumnIndex],
            amount: r[amountColumnIndex],
            fullRow: r
        })));
        return null;
    }
    
    if (isNaN(labour) || labour === 0) {
        console.error('LABOUR is missing or zero after search');
        return null;
    }
    
    // Ensure noOfUnits is valid
    if (!noOfUnits || noOfUnits <= 0 || isNaN(noOfUnits)) {
        console.warn('Invalid NO OF UNITS:', noOfUnits);
        return null;
    }
    
    // If sumAfterLabour not found, calculate it
    if (sumAfterLabour === 0 || isNaN(sumAfterLabour)) {
        sumAfterLabour = total + labour;
    }
    
    // Apply formula as per user requirements
    // SUM AFTER NUMBER OF UNITS = (TOTAL + LABOUR) * NO OF UNITS
    const sumAfterNumberOfUnits = (total + labour) * noOfUnits;
    
    // Validate calculation
    if (isNaN(sumAfterNumberOfUnits) || !isFinite(sumAfterNumberOfUnits)) {
        console.error('Error calculating SUM AFTER NUMBER OF UNITS:', { total, labour, noOfUnits });
        return null;
    }
    
    // OVER HEAD = SUM AFTER NUMBER OF UNITS * 50%
    const overhead = sumAfterNumberOfUnits * 0.50;
    
    // SUM AFTER OVERHEAD = OVER HEAD + SUM AFTER NUMBER OF UNITS
    const sumAfterOverhead = overhead + sumAfterNumberOfUnits;
    
    // TAX 5% = SUM AFTER OVERHEAD * 5%
    const tax5 = sumAfterOverhead * 0.05;
    
    // SUM AFTER TAX = SUM AFTER OVERHEAD + TAX 5%
    const sumAfterTax = sumAfterOverhead + tax5;
    
    // PROVISIONAL SUM 10% = SUM AFTER TAX * 10%
    const provisionalSum10 = sumAfterTax * 0.10;
    
    // NET TOTAL = SUM AFTER TAX + PROVISIONAL SUM 10%
    const netTotal = sumAfterTax + provisionalSum10;
    
    // Debug logging
    console.log('DB Calculation:', {
        total,
        labour,
        sumAfterLabour,
        noOfUnits,
        sumAfterNumberOfUnits,
        overhead,
        sumAfterOverhead,
        tax5,
        sumAfterTax,
        provisionalSum10,
        netTotal
    });
    
    return {
        total,
        labour,
        sumAfterLabour,
        noOfUnits,
        sumAfterNumberOfUnits,
        overhead,
        sumAfterOverhead,
        tax5,
        sumAfterTax,
        provisionalSum10,
        netTotal
    };
}

// Display detailed table with recalculated values for special DBs
function displayDetailedTableWithRecalculation(data, container, sheetName, calculation) {
    if (!data || data.length === 0) {
        container.innerHTML = '<div class="error-message">No data found in sheet "' + sheetName + '"</div>';
        return;
    }
    
    // First row as headers
    const headers = data[0] || [];
    const rows = data.slice(1);
    
    // Find column indices
    const itemColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('item') || 
              h.toString().toLowerCase().includes('description'))
    );
    
    // Find AMOUNT column specifically (prioritize exact match, then partial)
    let amountColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().trim() === 'amount'
    );
    
    // If exact match not found, try partial match
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('amount') && 
                  !h.toString().toLowerCase().includes('price'))
        );
    }
    
    // Fallback to cost/estimate if amount not found
    if (amountColumnIndex === -1) {
        amountColumnIndex = headers.findIndex(h => 
            h && (h.toString().toLowerCase().includes('cost') ||
                  h.toString().toLowerCase().includes('estimate'))
        );
    }
    
    console.log('Column indices:', { itemColumnIndex, amountColumnIndex });
    console.log('Headers:', headers);
    console.log('Calculation values:', calculation);
    
    if (itemColumnIndex === -1 || amountColumnIndex === -1) {
        container.innerHTML = '<div class="error-message">Could not find required columns in sheet</div>';
        return;
    }
    
    // List of calculation-related row names that should always be shown
    const calculationRows = [
        'no of units',
        'number of units',
        'sum after number of units',
        'sum after no of units',
        'over head',
        'overhead',
        'sum after overhead',
        'tax',
        'tax 5%',
        'sum after tax',
        'provisional sum',
        'provisional sum 10%',
        'net total'
    ];
    
    // Debug: Log all rows to see what we're working with
    console.log('Sample rows for matching:', rows.slice(0, 30).map((row, idx) => ({
        index: idx,
        item: row[itemColumnIndex],
        amount: row[amountColumnIndex],
        fullRow: row
    })));
    
    // Find rows to modify and ensure calculation rows are included
    const modifiedRows = rows.map((row, rowIndex) => {
        const itemValue = row[itemColumnIndex];
        if (!itemValue) return row;
        
        const itemStr = itemValue.toString().toLowerCase().trim();
        const newRow = [...row];
        let updated = false;
        let matchedLabel = '';
        
        // Update specific rows based on formula - use more flexible matching
        // Check for "NO OF UNITS" or "NUMBER OF UNITS" (but not "SUM AFTER NUMBER OF UNITS")
        if ((itemStr === 'no of units' || itemStr === 'number of units') || 
            (itemStr.includes('no') && itemStr.includes('units') && !itemStr.includes('sum after') && !itemStr.includes('number'))) {
            newRow[amountColumnIndex] = calculation.noOfUnits;
            updated = true;
            matchedLabel = 'NO OF UNITS';
            console.log(`✓ Updated NO OF UNITS row ${rowIndex}: "${itemValue}" -> ${calculation.noOfUnits}`);
        } 
        // Check for "SUM AFTER NUMBER OF UNITS"
        else if (itemStr.includes('sum after number of units') || 
                 itemStr.includes('sum after no of units') ||
                 (itemStr.includes('sum after') && itemStr.includes('number') && itemStr.includes('units'))) {
            newRow[amountColumnIndex] = calculation.sumAfterNumberOfUnits;
            updated = true;
            matchedLabel = 'SUM AFTER NUMBER OF UNITS';
            console.log(`✓ Updated SUM AFTER NUMBER OF UNITS row ${rowIndex}: "${itemValue}" -> ${calculation.sumAfterNumberOfUnits}`);
        } 
        // Check for "OVER HEAD" or "OVERHEAD"
        else if (itemStr === 'over head' || itemStr === 'overhead' || 
                 (itemStr.includes('over') && itemStr.includes('head') && !itemStr.includes('sum after'))) {
            newRow[amountColumnIndex] = calculation.overhead;
            updated = true;
            matchedLabel = 'OVER HEAD';
            console.log(`✓ Updated OVER HEAD row ${rowIndex}: "${itemValue}" -> ${calculation.overhead}`);
        } 
        // Check for "SUM AFTER OVERHEAD"
        else if (itemStr.includes('sum after overhead')) {
            newRow[amountColumnIndex] = calculation.sumAfterOverhead;
            updated = true;
            matchedLabel = 'SUM AFTER OVERHEAD';
            console.log(`✓ Updated SUM AFTER OVERHEAD row ${rowIndex}: "${itemValue}" -> ${calculation.sumAfterOverhead}`);
        } 
        // Check for "TAX 5%" (but not "SUM AFTER TAX")
        else if ((itemStr.includes('tax') && itemStr.includes('5%')) || 
                 (itemStr === 'tax 5%') || 
                 (itemStr === 'tax' && !itemStr.includes('sum after'))) {
            newRow[amountColumnIndex] = calculation.tax5;
            updated = true;
            matchedLabel = 'TAX 5%';
            console.log(`✓ Updated TAX 5% row ${rowIndex}: "${itemValue}" -> ${calculation.tax5}`);
        } 
        // Check for "SUM AFTER TAX"
        else if (itemStr.includes('sum after tax')) {
            newRow[amountColumnIndex] = calculation.sumAfterTax;
            updated = true;
            matchedLabel = 'SUM AFTER TAX';
            console.log(`✓ Updated SUM AFTER TAX row ${rowIndex}: "${itemValue}" -> ${calculation.sumAfterTax}`);
        } 
        // Check for "PROVISIONAL SUM 10%"
        else if ((itemStr.includes('provisional sum') && itemStr.includes('10%')) || 
                 itemStr.includes('provisional sum 10%') ||
                 (itemStr.includes('provisional') && itemStr.includes('sum') && !itemStr.includes('after'))) {
            newRow[amountColumnIndex] = calculation.provisionalSum10;
            updated = true;
            matchedLabel = 'PROVISIONAL SUM 10%';
            console.log(`✓ Updated PROVISIONAL SUM 10% row ${rowIndex}: "${itemValue}" -> ${calculation.provisionalSum10}`);
        } 
        // Check for "NET TOTAL"
        else if (itemStr.includes('net total')) {
            newRow[amountColumnIndex] = calculation.netTotal;
            updated = true;
            matchedLabel = 'NET TOTAL';
            console.log(`✓ Updated NET TOTAL row ${rowIndex}: "${itemValue}" -> ${calculation.netTotal}`);
        }
        
        if (updated) {
            console.log(`Row ${rowIndex} before:`, row[amountColumnIndex], `after:`, newRow[amountColumnIndex]);
        }
        
        return newRow;
    });
    
    // Count how many rows were updated
    const updatedCount = modifiedRows.filter((row, idx) => {
        const itemValue = row[itemColumnIndex];
        if (!itemValue) return false;
        const itemStr = itemValue.toString().toLowerCase().trim();
        return itemStr.includes('no of units') || 
               itemStr.includes('sum after number') ||
               itemStr.includes('over head') ||
               itemStr.includes('sum after overhead') ||
               (itemStr.includes('tax') && !itemStr.includes('sum after')) ||
               itemStr.includes('sum after tax') ||
               itemStr.includes('provisional') ||
               itemStr.includes('net total');
    }).length;
    
    console.log(`Total calculation rows found and updated: ${updatedCount}`);
    
    // Display with modified rows, but use a flag to preserve calculation rows
    displayDetailedTable(modifiedRows, container, sheetName, headers, true);
}

// Display detailed data as table
function displayDetailedTable(data, container, sheetName, providedHeaders = null, preserveCalculationRows = false) {
    if (!data || data.length === 0) {
        container.innerHTML = '<div class="error-message">No data found in sheet "' + sheetName + '"</div>';
        return;
    }
    
    // First row as headers (or use provided headers)
    const headers = providedHeaders || data[0] || [];
    const rows = providedHeaders ? data : data.slice(1);
    
    // Find the "Amount" column index (case-insensitive)
    const amountColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('amount') || 
              h.toString().toLowerCase().includes('estimate') ||
              h.toString().toLowerCase().includes('cost'))
    );
    
    // Find item/description column index for checking calculation rows
    const itemColumnIndex = headers.findIndex(h => 
        h && (h.toString().toLowerCase().includes('item') || 
              h.toString().toLowerCase().includes('description'))
    );
    
    // Find BRAND column index
    const brandColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().includes('brand')
    );
    
    // Find PRICE column index
    const priceColumnIndex = headers.findIndex(h => 
        h && h.toString().toLowerCase().includes('price')
    );
    
    // Find price and amount column indices for formatting
    const priceAmountColumnIndices = headers.map((h, index) => {
        if (!h) return null;
        const headerLower = h.toString().toLowerCase();
        if (headerLower.includes('price') || 
            headerLower.includes('amount') || 
            headerLower.includes('cost') ||
            headerLower.includes('estimate') ||
            headerLower.includes('total')) {
            return index;
        }
        return null;
    }).filter(index => index !== null);
    
    // List of calculation-related row names that should always be shown
    const calculationRowKeywords = [
        'no of units',
        'number of units',
        'sum after number of units',
        'sum after no of units',
        'over head',
        'overhead',
        'sum after overhead',
        'tax 5%',
        'tax',
        'sum after tax',
        'provisional sum',
        'provisional sum 10%',
        'net total'
    ];
    
    // Filter rows: only show rows where amount column is not zero
    // BUT preserve calculation rows if preserveCalculationRows flag is set
    const filteredRows = rows.filter(row => {
        if (amountColumnIndex === -1) {
            // If no amount column found, show all rows
            return true;
        }
        
        // Check if this is a calculation row that should be preserved
        if (preserveCalculationRows && itemColumnIndex !== -1) {
            const itemValue = row[itemColumnIndex];
            if (itemValue) {
                const itemStr = itemValue.toString().toLowerCase().trim();
                const isCalculationRow = calculationRowKeywords.some(keyword => 
                    itemStr.includes(keyword)
                );
                if (isCalculationRow) {
                    return true; // Always show calculation rows
                }
            }
        }
        
        const amountValue = row[amountColumnIndex];
        // Check if amount is not zero (handle various formats: 0, "0", 0.0, null, undefined, empty string)
        if (amountValue === null || amountValue === undefined || amountValue === '') {
            return false;
        }
        const numValue = parseFloat(amountValue);
        return !isNaN(numValue) && numValue !== 0;
    });
    
    if (filteredRows.length === 0) {
        container.innerHTML = '<div class="error-message">No rows with non-zero amounts found in sheet "' + sheetName + '"</div>';
        return;
    }
    
    // Helper function to format numeric values
    const formatNumericValue = (value, columnIndex) => {
        if (priceAmountColumnIndices.includes(columnIndex)) {
            // Try to parse as number
            const numValue = parseFloat(value);
            if (!isNaN(numValue) && value !== null && value !== undefined && value !== '') {
                return numValue.toFixed(2);
            }
        }
        return value;
    };
    
    // Find indices of rows that should be highlighted (between "Total" and "Net Total")
    // Check all columns, not just the first one
    let highlightStartIndex = -1;
    let highlightEndIndex = -1;
    
    filteredRows.forEach((row, rowIndex) => {
        // Check all cells in the row for "TOTAL" or "NET TOTAL"
        for (let colIndex = 0; colIndex < row.length; colIndex++) {
            const cellValue = row[colIndex];
            if (cellValue !== null && cellValue !== undefined && cellValue !== '') {
                const cellValueStr = cellValue.toString().toLowerCase().trim();
                
                // Find "TOTAL" row - exact match or starts with "total" but not "net total" or "sum after"
                if (cellValueStr === 'total' || 
                    (cellValueStr.startsWith('total') && 
                     !cellValueStr.includes('net') && 
                     !cellValueStr.includes('sum after') &&
                     !cellValueStr.includes('after'))) {
                    if (highlightStartIndex === -1) {
                        highlightStartIndex = rowIndex;
                        console.log(`Found TOTAL at row ${rowIndex}:`, cellValue);
                        break; // Found it, move to next row
                    }
                }
                
                // Find "NET TOTAL" row
                if ((cellValueStr.includes('net') && cellValueStr.includes('total')) ||
                    cellValueStr === 'net total') {
                    highlightEndIndex = rowIndex;
                    console.log(`Found NET TOTAL at row ${rowIndex}:`, cellValue);
                    break; // Found it, move to next row
                }
            }
        }
    });
    
    // Debug logging
    console.log('=== HIGHLIGHT DEBUG ===');
    console.log('Highlight range:', highlightStartIndex, 'to', highlightEndIndex);
    console.log('Total filtered rows:', filteredRows.length);
    if (highlightStartIndex !== -1 && highlightEndIndex !== -1) {
        console.log('Will highlight rows:', highlightStartIndex, 'through', highlightEndIndex);
        filteredRows.slice(highlightStartIndex, highlightEndIndex + 1).forEach((row, idx) => {
            console.log(`Row ${highlightStartIndex + idx}:`, row);
        });
    } else {
        console.log('WARNING: Could not find TOTAL or NET TOTAL rows!');
        console.log('First 5 rows:', filteredRows.slice(0, 5).map(r => r[0]));
    }
    
    let html = `<p><strong>Sheet:</strong> ${sheetName}</p>`;
    html += '<table class="detailed-data-table"><thead><tr>';
    
    headers.forEach(header => {
        html += `<th>${header || ''}</th>`;
    });
    html += '</tr></thead><tbody>';
    
    filteredRows.forEach((row, rowIndex) => {
        // Determine if this row should be highlighted
        const shouldHighlight = highlightStartIndex !== -1 && 
                               highlightEndIndex !== -1 && 
                               rowIndex >= highlightStartIndex && 
                               rowIndex <= highlightEndIndex;
        
        if (shouldHighlight) {
            console.log(`Highlighting row ${rowIndex}:`, row[0]);
        }
        
        const rowClass = shouldHighlight ? 'highlight-row' : '';
        const bgColor = shouldHighlight ? '#ffeb3b' : '';
        const rowStyle = shouldHighlight ? ` style="background-color: ${bgColor} !important; font-weight: 600 !important;"` : '';
        html += `<tr class="${rowClass}"${rowStyle}>`;
        headers.forEach((_, index) => {
            let value = row[index] !== undefined ? row[index] : '';
            
            // Correct typo: "SUM AFTER LABOUT" -> "SUM AFTER LABOUR"
            if (typeof value === 'string') {
                value = value.replace(/SUM AFTER LABOUT/gi, 'SUM AFTER LABOUR');
            }
            
            // Hide '0' values in BRAND column
            if (index === brandColumnIndex) {
                if (value === '0' || value === 0 || value === '0.00' || value === null || value === undefined || value === '') {
                    value = ''; // Hide the '0' value
                }
            }
            
            // Hide '0' values in PRICE column
            if (index === priceColumnIndex) {
                const priceNum = value !== null && value !== undefined && value !== '' ? parseFloat(value) : 0;
                if (isNaN(priceNum) || priceNum === 0) {
                    value = ''; // Hide the '0' value
                }
            }
            
            const formattedValue = formatNumericValue(value, index);
            const cellBgColor = shouldHighlight ? '#ffeb3b' : '';
            const cellBorder = shouldHighlight ? '2px solid #ff9800' : '';
            const cellStyle = shouldHighlight ? ` style="background-color: ${cellBgColor} !important; border: ${cellBorder} !important; font-weight: 600 !important;"` : '';
            html += `<td${cellStyle}>${formattedValue}</td>`;
        });
        html += '</tr>';
    });
    
    html += '</tbody></table>';
    container.innerHTML = html;
}

// Close detailed page and return to diagram
function closeDetailedPage() {
    const detailedPage = document.getElementById('detailed-page');
    const sourceView = detailedPage.getAttribute('data-source-view') || 'main';
    
    detailedPage.classList.add('hidden');
    
    // Return to the appropriate view based on source
    if (sourceView === 'mdb') {
        // Return to MDB view
        document.getElementById('mdb-view-page').classList.remove('hidden');
    } else {
        // Return to main content
        document.querySelector('main > .controls').style.display = 'block';
        document.querySelector('main > .legend-panel').style.display = 'block';
        document.querySelector('main > .info-panel').style.display = 'block';
        document.getElementById('main-content').style.display = 'flex';
    }
}

// Event listeners
document.getElementById('closeDetails').addEventListener('click', () => {
    document.getElementById('details-panel').classList.add('hidden');
});

document.getElementById('backToHome').addEventListener('click', () => {
    closeDetailedPage();
});

document.getElementById('searchBtn').addEventListener('click', searchItems);
document.getElementById('searchInput').addEventListener('keypress', (e) => {
    if (e.key === 'Enter') {
        searchItems();
    }
});

document.getElementById('kindFilter').addEventListener('change', filterByKind);

document.getElementById('resetView').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().duration(750).call(
            zoom.transform,
            d3.zoomIdentity
        );
    }
});

document.getElementById('fitView').addEventListener('click', () => {
    if (svg && g) {
        const bounds = g.node().getBBox();
        const fullWidth = svg.node().getBoundingClientRect().width;
        const fullHeight = svg.node().getBoundingClientRect().height;
        const widthScale = (fullWidth - 40) / bounds.width;
        const heightScale = (fullHeight - 40) / bounds.height;
        const scale = Math.min(widthScale, heightScale, 1);
        
        const translate = [
            (fullWidth - bounds.width * scale) / 2 - bounds.x * scale,
            (fullHeight - bounds.height * scale) / 2 - bounds.y * scale
        ];
        
        svg.transition().duration(750).call(
            zoom.transform,
            d3.zoomIdentity.translate(translate[0], translate[1]).scale(scale)
        );
    }
});

document.getElementById('expandAll').addEventListener('click', () => {
    if (root) {
        root.each(d => {
            if (d._children) {
                d.children = d._children;
                d._children = null;
            }
        });
        initializeDiagram();
    }
});

document.getElementById('collapseAll').addEventListener('click', () => {
    if (root) {
        root.children.forEach(collapse);
        root.children.forEach(collapse);
        root.children.forEach(collapse);
        initializeDiagram();
    }
});

function collapse(d) {
    if (d.children) {
        d._children = d.children;
        d._children.forEach(collapse);
        d.children = null;
    }
}

document.getElementById('zoomIn').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().call(zoom.scaleBy, 1.3);
    }
});

document.getElementById('zoomOut').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().call(zoom.scaleBy, 0.77);
    }
});

document.getElementById('zoomReset').addEventListener('click', () => {
    if (svg && zoom) {
        svg.transition().duration(750).call(
            zoom.transform,
            d3.zoomIdentity
        );
    }
});

document.getElementById('toggleTreeView').addEventListener('click', () => {
    const sidebar = document.getElementById('tree-sidebar');
    sidebar.classList.toggle('hidden');
});

document.getElementById('closeTreeView').addEventListener('click', () => {
    document.getElementById('tree-sidebar').classList.add('hidden');
});

// MDB View button
document.getElementById('showMDBView').addEventListener('click', () => {
    openMDBView();
});

// Back button from MDB view
document.getElementById('backToHomeFromMDB').addEventListener('click', () => {
    closeMDBView();
});

// MDB Tab buttons
document.querySelectorAll('.mdb-tab').forEach(tab => {
    tab.addEventListener('click', () => {
        // Remove active class from all tabs
        document.querySelectorAll('.mdb-tab').forEach(t => t.classList.remove('active'));
        // Add active class to clicked tab
        tab.classList.add('active');
        // Load the corresponding MDB table
        const mdbName = tab.getAttribute('data-mdb');
        loadMDBTable(mdbName);
    });
});

function searchItems() {
    const searchTerm = document.getElementById('searchInput').value.toLowerCase().trim();
    if (!searchTerm) {
        resetFilter();
        return;
    }
    
    // Highlight matching nodes
    if (svg && g) {
        g.selectAll('g.node').each(function(d) {
            const node = d3.select(this);
            const name = (d.data.name || '').toLowerCase();
            if (name.includes(searchTerm)) {
                node.select('circle').attr('stroke-width', 4).attr('stroke', '#FFD700');
            } else {
                node.select('circle').attr('stroke-width', d.data.id === 'RMU' ? 3 : 2)
                    .attr('stroke', getNodeBorderColor(d.data.kind));
            }
        });
    }
}

function filterByKind() {
    const selectedKind = document.getElementById('kindFilter').value;
    if (!selectedKind) {
        resetFilter();
        return;
    }
    
    if (svg && g) {
        g.selectAll('g.node').each(function(d) {
            const node = d3.select(this);
            if (d.data.kind === selectedKind) {
                node.style('opacity', 1);
            } else {
                node.style('opacity', 0.2);
            }
        });
    }
}

function resetFilter() {
    if (svg && g) {
        g.selectAll('g.node').each(function(d) {
            const node = d3.select(this);
            node.style('opacity', 1);
            node.select('circle').attr('stroke-width', d.data.id === 'ROOT' ? 3 : 2)
                .attr('stroke', getNodeBorderColor(d.data.kind));
        });
    }
}

// Handle window resize
window.addEventListener('resize', () => {
    if (svg) {
        const container = d3.select('#mindmap');
        const width = container.node().offsetWidth || 1200;
        const height = Math.max(600, window.innerHeight * 0.6);
        svg.attr('width', width).attr('height', height);
    }
});

// MDB View Functions
function openMDBView() {
    document.getElementById('main-content').style.display = 'none';
    document.getElementById('detailed-page').classList.add('hidden');
    document.getElementById('mdb-view-page').classList.remove('hidden');
    
    // Load MDB1 by default
    loadMDBTable('MDB1');
}

function closeMDBView() {
    document.getElementById('mdb-view-page').classList.add('hidden');
    document.getElementById('main-content').style.display = 'block';
}

function loadMDBTable(mdbName) {
    if (!treeData) {
        console.error('Tree data not available');
        return;
    }
    
    // Find the MDB node
    const mdbNode = treeData.children.find(child => child.name === mdbName);
    if (!mdbNode) {
        document.getElementById('mdb-table-container').innerHTML = 
            `<div class="error-message">No data found for ${mdbName}</div>`;
        return;
    }
    
    // Show loading message
    document.getElementById('mdb-table-container').innerHTML = 
        '<div class="loading">Loading and calculating estimates...</div>';
    
    // Flatten the tree structure into a table
    const tableData = flattenTreeForTable(mdbNode);
    
    // Recalculate estimates for special DBs
    recalculateMDBTableEstimates(tableData).then(() => {
        // Generate HTML table with recalculated estimates
        const html = generateMDBTable(tableData, mdbName);
        document.getElementById('mdb-table-container').innerHTML = html;
        
        // Setup collapse functionality after HTML is inserted
        setTimeout(() => {
            setupMDBTableCollapse(mdbName);
        }, 100);
    }).catch(error => {
        console.error('Error recalculating estimates:', error);
        // Fallback to table without recalculation
        const html = generateMDBTable(tableData, mdbName);
        document.getElementById('mdb-table-container').innerHTML = html;
        
        // Setup collapse functionality after HTML is inserted
        setTimeout(() => {
            setupMDBTableCollapse(mdbName);
        }, 100);
    });
}

// Recalculate estimates for special DBs in MDB table
async function recalculateMDBTableEstimates(tableData) {
    const specialDBs = ['DB.TN.LXX.1B1.01', 'DB.TN.LXX.2B1.01', 'DB.TN.LXX.3B1.01'];
    
    // Load Excel file once
    const response = await fetch('e2.xlsx');
    if (!response.ok) {
        console.warn('Could not load Excel file for recalculation');
        return;
    }
    const arrayBuffer = await response.arrayBuffer();
    const workbook = XLSX.read(arrayBuffer, { type: 'array' });
    
    // Process each row to find special DBs and recalculate
    for (let i = 0; i < tableData.length; i++) {
        const row = tableData[i];
        const itemName = row.name || row.itemdrop || '';
        
        // Check if this is a special DB (exact match or group node)
        let dbName = null;
        let noOfUnits = 1;
        
        // Check for exact match
        if (specialDBs.includes(itemName)) {
            dbName = itemName;
        } 
        // Check for group node pattern: "DB.TN.LXX.1B1.01 (5 copies)"
        else {
            for (const dbType of specialDBs) {
                if (itemName.includes(dbType) && itemName.includes('copies')) {
                    dbName = dbType;
                    // Extract copy count from name like "DB.TN.LXX.1B1.01 (5 copies)"
                    const match = itemName.match(/\((\d+)\s*copies?\)/i);
                    if (match) {
                        noOfUnits = parseInt(match[1]) || 1;
                    }
                    break;
                }
            }
        }
        
        if (dbName) {
            // Find parent SMDB from fedFrom
            const fedFrom = row.fedFrom || '';
            const smdbList = fedFrom.split('\n').map(p => p.trim()).filter(p => p && p !== 'RMU' && !p.startsWith('MDB'));
            
            // If noOfUnits not extracted from group name, get from copy counts
            if (noOfUnits === 1 && smdbList.length > 0) {
                const smdbName = smdbList[0];
                const copyCounts = getDBCopyCounts(smdbName);
                noOfUnits = copyCounts && copyCounts[dbName] ? copyCounts[dbName] : 1;
            }
            
            if (smdbList.length > 0 || noOfUnits > 1) {
                // Calculate estimate for this DB
                const calculation = await calculateDBEstimateFromSheet(workbook, dbName, noOfUnits);
                
                if (calculation && calculation.netTotal) {
                    // Update the estimate with recalculated NET TOTAL
                    row.estimate = calculation.netTotal;
                    row.recalculated = true;
                    console.log(`Recalculated ${itemName} (${dbName}, ${noOfUnits} units) estimate: ${calculation.netTotal}`);
                }
            }
        }
    }
}

// Calculate DB estimate from Excel workbook
async function calculateDBEstimateFromSheet(workbook, dbName, noOfUnits) {
    const sheetNames = workbook.SheetNames;
    
    // Find DB sheet
    let dbSheetName = sheetNames.find(name => 
        name === dbName || 
        name.toLowerCase() === dbName.toLowerCase()
    );
    
    if (!dbSheetName) {
        dbSheetName = sheetNames.find(name => 
            name.toLowerCase().includes(dbName.toLowerCase()) ||
            dbName.toLowerCase().includes(name.toLowerCase())
        );
    }
    
    if (!dbSheetName) {
        console.warn(`Sheet not found for ${dbName}`);
        return null;
    }
    
    const dbWorksheet = workbook.Sheets[dbSheetName];
    const dbData = XLSX.utils.sheet_to_json(dbWorksheet, { header: 1 });
    
    // Calculate estimate for this DB
    return calculateDBEstimate(dbData, noOfUnits);
}

function flattenTreeForTable(node, level = 0, parentName = '') {
    const rows = [];
    
    // Add current node - ensure all properties are populated from data if available
    const nodeData = node.data || {};
    const row = {
        level: level,
        name: node.name,
        kind: node.kind || nodeData.KIND || 'Unknown',
        load: node.load || nodeData.Load || '0 kW',
        estimate: node.estimate !== undefined ? node.estimate : (nodeData.Estimate || 0),
        fedFrom: parentName || nodeData['FED FROM'] || '',
        noOfItems: nodeData['NO OF ITEMS'] !== undefined ? nodeData['NO OF ITEMS'] : '',
        itemdrop: nodeData.Itemdrop || node.name || '',
        mdb: node.mdb || nodeData.MDB || '',
        data: nodeData
    };
    rows.push(row);
    
    // Process children
    if (node.children && node.children.length > 0) {
        node.children.forEach(child => {
            const childRows = flattenTreeForTable(child, level + 1, node.name);
            rows.push(...childRows);
        });
    }
    
    return rows;
}

function generateMDBTable(data, mdbName) {
    if (!data || data.length === 0) {
        return '<div class="error-message">No data available</div>';
    }
    
    // Group data by parent to enable collapsing
    const groupedData = groupDataByParent(data);
    
    let html = `<h3 style="margin-bottom: 20px; color: #667eea;">${mdbName} - ${data.length} items</h3>`;
    html += '<table class="mdb-data-table" id="mdb-table-' + mdbName + '">';
    html += '<thead><tr>';
    html += '<th>Name</th>';
    html += '<th>Kind</th>';
    html += '<th>Load</th>';
    html += '<th>No. of Items</th>';
    html += '<th>Fed From</th>';
    html += '<th>Estimate (AED)</th>';
    html += '</tr></thead><tbody>';
    
    // Generate rows with collapsible functionality
    groupedData.forEach((group, groupIndex) => {
        const parentRow = group.parent;
        const children = group.children;
        const hasChildren = children.length > 0;
        const indent = '&nbsp;'.repeat(parentRow.level * 4);
        const kindClass = `kind-${(parentRow.kind || '').toLowerCase().replace(/\s+/g, '-')}`;
        const estimate = parseFloat(parentRow.estimate) || 0;
        const isRecalculated = parentRow.recalculated || false;
        const isParentType = parentRow.kind && (
            parentRow.kind.toLowerCase().includes('smdb') || 
            parentRow.kind.toLowerCase().includes('esmdb') || 
            parentRow.kind.toLowerCase().includes('mcc')
        );
        const rowId = `mdb-row-${mdbName}-${groupIndex}`;
        const collapseId = `collapse-${mdbName}-${groupIndex}`;
        
        // Parent row (SMDB, ESMDB, MCC or other parent)
        const isClickable = estimate > 0;
        const clickableClass = isClickable ? 'clickable-row' : 'non-clickable-row';
        // Ensure data is properly serialized - use itemdrop or name to find original data if node.data is empty
        let rowData = parentRow.data || {};
        if (!rowData || Object.keys(rowData).length === 0) {
            // Try to find data from allData if available
            const itemName = parentRow.name || parentRow.itemdrop || '';
            if (itemName && window.allData) {
                const foundData = window.allData.find(d => (d.Itemdrop || d.MDB) === itemName);
                if (foundData) {
                    rowData = foundData;
                }
            }
        }
        html += `<tr class="mdb-parent-row ${hasChildren && isParentType ? 'has-children' : ''} ${clickableClass}" data-row-id="${rowId}" data-collapse-id="${collapseId}" data-item-name="${parentRow.name || ''}" data-item-data='${JSON.stringify(rowData)}' data-estimate="${estimate}">`;
        html += `<td>${indent}`;
        if (hasChildren && isParentType) {
            html += `<span class="collapse-toggle" style="cursor: pointer; margin-right: 5px; user-select: none;">▶</span>`;
        }
        html += `${parentRow.name || ''}${isRecalculated ? ' <span style="color: #4CAF50; font-size: 0.85em;">(recalculated)</span>' : ''}</td>`;
        html += `<td><span class="${kindClass}">${parentRow.kind || ''}</span></td>`;
        html += `<td>${parentRow.load || '0 kW'}</td>`;
        html += `<td>${parentRow.noOfItems || ''}</td>`;
        html += `<td>${parentRow.fedFrom || ''}</td>`;
        html += `<td class="numeric"${isRecalculated ? ' style="font-weight: 600; color: #4CAF50;"' : ''}>${formatNumber(estimate)}</td>`;
        html += '</tr>';
        
        // Child rows - initially hidden if parent is SMDB, ESMDB, or MCC
        if (hasChildren && isParentType) {
            children.forEach((childRow, childIndex) => {
                const childIndent = '&nbsp;'.repeat((childRow.level - parentRow.level) * 4 + 8); // Extra indent for children
                const childKindClass = `kind-${(childRow.kind || '').toLowerCase().replace(/\s+/g, '-')}`;
                const childEstimate = parseFloat(childRow.estimate) || 0;
                const childIsRecalculated = childRow.recalculated || false;
                const childRowClass = `mdb-child-row ${collapseId}`;
                const childIsClickable = childEstimate > 0;
                const childClickableClass = childIsClickable ? 'clickable-row' : 'non-clickable-row';
                // Ensure data is properly serialized - use itemdrop or name to find original data if node.data is empty
                let childRowData = childRow.data || {};
                if (!childRowData || Object.keys(childRowData).length === 0) {
                    // Try to find data from allData if available
                    const childItemName = childRow.name || childRow.itemdrop || '';
                    if (childItemName && window.allData) {
                        const foundData = window.allData.find(d => (d.Itemdrop || d.MDB) === childItemName);
                        if (foundData) {
                            childRowData = foundData;
                        }
                    }
                }
                html += `<tr class="${childRowClass} ${childClickableClass}" style="display: none;" data-item-name="${childRow.name || ''}" data-item-data='${JSON.stringify(childRowData)}' data-estimate="${childEstimate}">`;
                html += `<td>${childIndent}${childRow.name || ''}${childIsRecalculated ? ' <span style="color: #4CAF50; font-size: 0.85em;">(recalculated)</span>' : ''}</td>`;
                html += `<td><span class="${childKindClass}">${childRow.kind || ''}</span></td>`;
                html += `<td>${childRow.load || '0 kW'}</td>`;
                html += `<td>${childRow.noOfItems || ''}</td>`;
                html += `<td>${childRow.fedFrom || ''}</td>`;
                html += `<td class="numeric"${childIsRecalculated ? ' style="font-weight: 600; color: #4CAF50;"' : ''}>${formatNumber(childEstimate)}</td>`;
                html += '</tr>';
            });
        }
    });
    
    html += '</tbody></table>';
    
    // Calculate totals - sum of MDB and all its descendants
    // The data array already includes the MDB node (first row) and all its children
    const totalEstimate = data.reduce((sum, row) => {
        const estimate = parseFloat(row.estimate) || 0;
        return sum + estimate;
    }, 0);
    
    const totalLoad = data.reduce((sum, row) => {
        const loadStr = row.load || '0 kW';
        const loadValue = parseFloat(loadStr.replace(/[^\d.]/g, '')) || 0;
        return sum + loadValue;
    }, 0);
    
    // Display totals prominently
    html += '<div style="margin-top: 20px; padding: 20px; background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); border-radius: 8px; color: white; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">';
    html += `<div style="font-size: 1.1em; margin-bottom: 10px;"><strong>${mdbName} Summary</strong></div>`;
    html += `<div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 15px;">`;
    html += `<div><strong>Total Load:</strong> ${totalLoad.toFixed(2)} kW</div>`;
    html += `<div style="font-size: 1.2em; font-weight: bold;"><strong>Total Estimate:</strong> ${formatNumber(totalEstimate)} AED</div>`;
    html += `</div>`;
    html += `<div style="margin-top: 10px; font-size: 0.9em; opacity: 0.9;">Includes ${mdbName} and all boards/equipment under it</div>`;
    html += '</div>';
    
    return html;
}

// Setup collapse functionality for MDB table
function setupMDBTableCollapse(mdbName) {
    const table = document.getElementById('mdb-table-' + mdbName);
    if (!table) {
        console.warn('Table not found:', 'mdb-table-' + mdbName);
        return;
    }
    
    const parentRows = table.querySelectorAll('.mdb-parent-row.has-children');
    console.log('Found parent rows with children:', parentRows.length);
    
    parentRows.forEach((parentRow, index) => {
        const collapseId = parentRow.getAttribute('data-collapse-id');
        const toggle = parentRow.querySelector('.collapse-toggle');
        
        if (!collapseId || !toggle) {
            console.warn('Missing collapseId or toggle for row', index);
            return;
        }
        
        // Query child rows using the collapseId class
        const childRows = Array.from(table.querySelectorAll('tr')).filter(row => 
            row.classList.contains('mdb-child-row') && row.classList.contains(collapseId)
        );
        
        console.log(`Row ${index}: collapseId=${collapseId}, childRows=${childRows.length}`);
        
        if (childRows.length > 0) {
            // Collapse handler - only for toggle icon
            const collapseHandler = function(e) {
                e.preventDefault();
                e.stopPropagation();
                
                const isCollapsed = childRows[0].style.display === 'none' || 
                                   childRows[0].style.display === '' ||
                                   window.getComputedStyle(childRows[0]).display === 'none';
                
                childRows.forEach(row => {
                    row.style.display = isCollapsed ? 'table-row' : 'none';
                });
                
                toggle.textContent = isCollapsed ? '▼' : '▶';
                
                console.log(`Toggled ${collapseId}: ${isCollapsed ? 'expanded' : 'collapsed'}`);
            };
            
            // Only attach collapse handler to toggle icon
            toggle.addEventListener('click', collapseHandler);
        }
        
        // Add click handler for opening detailed page (for all parent rows with non-zero estimates)
        const itemName = parentRow.getAttribute('data-item-name');
        const itemDataStr = parentRow.getAttribute('data-item-data');
        const rowEstimate = parseFloat(parentRow.getAttribute('data-estimate')) || 0;
        
        if (itemName && itemDataStr && rowEstimate > 0) {
            parentRow.style.cursor = 'pointer';
            parentRow.setAttribute('data-click-handler', 'true');
            parentRow.addEventListener('click', function(e) {
                // Don't trigger if clicking on collapse toggle
                if (e.target.classList.contains('collapse-toggle') || 
                    e.target.closest('.collapse-toggle')) {
                    return;
                }
                
                try {
                    const itemData = JSON.parse(itemDataStr);
                    openDetailedPage(itemData, itemName, 'mdb');
                } catch (err) {
                    console.error('Error parsing item data:', err);
                    openDetailedPage(null, itemName, 'mdb');
                }
            });
        } else if (rowEstimate === 0) {
            // Make non-clickable rows visually distinct
            parentRow.style.cursor = 'not-allowed';
            parentRow.style.opacity = '0.6';
        }
    });
    
    // Add click handlers for ALL rows with non-zero estimates (including MDB, PFC, BusBarRaiser, etc.)
    const allRowsWithData = table.querySelectorAll('tr[data-item-name]');
    allRowsWithData.forEach(row => {
        const itemName = row.getAttribute('data-item-name');
        const itemDataStr = row.getAttribute('data-item-data');
        const rowEstimate = parseFloat(row.getAttribute('data-estimate')) || 0;
        
        // Skip if already has click handler (parent rows processed above)
        if (row.hasAttribute('data-click-handler')) {
            return;
        }
        
        // Only make clickable if estimate > 0 and has valid data
        if (itemName && itemDataStr && rowEstimate > 0) {
            row.style.cursor = 'pointer';
            row.setAttribute('data-click-handler', 'true');
            row.addEventListener('click', function(e) {
                // Don't trigger if clicking on collapse toggle
                if (e.target.classList.contains('collapse-toggle') || 
                    e.target.closest('.collapse-toggle')) {
                    return;
                }
                
                try {
                    const itemData = JSON.parse(itemDataStr);
                    console.log('Opening detailed page for:', itemName, 'Data:', itemData);
                    openDetailedPage(itemData, itemName, 'mdb');
                } catch (err) {
                    console.error('Error parsing item data:', err, 'Item:', itemName, 'Data string:', itemDataStr);
                    openDetailedPage(null, itemName, 'mdb');
                }
            });
        } else if (rowEstimate === 0) {
            // Make non-clickable rows visually distinct
            row.style.cursor = 'not-allowed';
            row.style.opacity = '0.6';
        } else {
            // Debug: log rows that should be clickable but aren't
            console.warn('Row not clickable:', {
                itemName,
                hasData: !!itemDataStr,
                estimate: rowEstimate,
                row: row
            });
        }
    });
}

// Group data by parent for collapsible display
function groupDataByParent(data) {
    const groups = [];
    
    for (let i = 0; i < data.length; i++) {
        const row = data[i];
        const isParentType = row.kind && (
            row.kind.toLowerCase().includes('smdb') || 
            row.kind.toLowerCase().includes('esmdb') || 
            row.kind.toLowerCase().includes('mcc')
        );
        
        // Check if this is an SMDB, ESMDB, or MCC that might have children
        if (isParentType) {
            // Look ahead to find children that belong to this parent
            const children = [];
            for (let j = i + 1; j < data.length; j++) {
                const nextRow = data[j];
                
                // If we hit another parent type or item at same or lower level, stop
                const nextIsParentType = nextRow.kind && (
                    nextRow.kind.toLowerCase().includes('smdb') || 
                    nextRow.kind.toLowerCase().includes('esmdb') || 
                    nextRow.kind.toLowerCase().includes('mcc')
                );
                
                if (nextRow.level <= row.level || nextIsParentType) {
                    break;
                }
                
                // If it's a direct child (one level deeper), add it
                if (nextRow.level === row.level + 1) {
                    children.push(nextRow);
                }
            }
            
            groups.push({
                parent: row,
                children: children
            });
            
            // Skip the children rows in the main loop (they're already added)
            i += children.length;
        } 
        // Standalone row (not a parent type with children)
        else {
            groups.push({
                parent: row,
                children: []
            });
        }
    }
    
    return groups;
}

// Note: loadData() is now called from the authentication handler after successful login
// The original DOMContentLoaded listener has been moved to the authentication section
