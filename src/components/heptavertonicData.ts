export const CANVAS_WIDTH = 1400;
export const CANVAS_HEIGHT = 1000;
export const TRIAD_SIZE = 180;
export const VERTEX_SIZE = 60;

export interface Vertex {
  id: string;
  name: string;
  shortLabel: string;
  function: string;
  description: string;
  cognitiveCenter: string;
  dimension: string;
  departments: string[];
  accountRange: string;
  position: {
    x: number;
    y: number;
  };
}

export interface Triad {
  id: string;
  name: string;
  shortLabel: string;
  accountRange: string;
  cognitiveCenter: string;
  dimension: string;
  description: string;
  departments: string[];
  financialSummary: string[];
  position: {
    x: number;
    y: number;
  };
  vertices: Vertex[];
}

export interface Flow {
  id: string;
  from: string;
  to: string;
  type: 'supply' | 'demand' | 'innovation' | 'complementary' | 'ai';
  label: string;
  bidirectional?: boolean;
}

export interface FinancialItem {
  name: string;
  value: number;
}

export interface FinancialCenter {
  name: string;
  total: number;
  items: FinancialItem[];
}

export interface FinancialData {
  centers: FinancialCenter[];
  eliminations: FinancialItem[];
  adjustments: FinancialItem[];
  transformations: FinancialItem[];
}

export interface PeripheralNode {
  id: string;
  name: string;
  shortLabel: string;
  type: 'core' | 'executive' | 'ai';
  description: string;
  accountRange?: string;
  dimension: string;
  departments: string[];
  financialSummary: string[];
  position: {
    x: number;
    y: number;
  };
}

export interface SimulationStep {
  id: string;
  title: string;
  description: string;
  activeNodeIds: string[];
  activeFlowIds: string[];
}

export const TRIADS: Triad[] = [
  {
    id: 'triad1',
    name: 'Triad 1 · Market Intelligence',
    shortLabel: 'τ₁',
    accountRange: '8000-8999',
    cognitiveCenter: 'Performance / Perception',
    dimension: 'Performance Dimension',
    description: 'Perceives market demand, pattern shifts, and institutional memory through customer, cash, and receivable signals.',
    departments: ['Market Intelligence Division', 'Competitive Analysis Unit', 'Strategic Memory Archive'],
    financialSummary: ['Market Intelligence Division: $3,200,000', 'Competitive Analysis Unit: $2,100,000', 'Strategic Memory Archive: $1,720,000'],
    position: { x: 300, y: 250 },
    vertices: [
      {
        id: 'v1a',
        name: 'T₁-OS · Perception',
        shortLabel: 'T₁-OS',
        function: 'Perception',
        description: 'Senses market conditions and the organization’s response capacity via customer and banking signals.',
        cognitiveCenter: 'System 4 · Perception',
        dimension: 'Performance Dimension',
        departments: ['Customer Control', 'Banking Positioning'],
        accountRange: '8000-8999',
        position: { x: -55, y: -48 },
      },
      {
        id: 'v1b',
        name: 'T₄-SA · Pattern',
        shortLabel: 'T₄-SA',
        function: 'Pattern',
        description: 'Recognizes repeating market and operational structures that inform adaptation.',
        cognitiveCenter: 'System 4 · Organization',
        dimension: 'Commitment Dimension',
        departments: ['Pattern Analysis', 'Competitive Signal Processing'],
        accountRange: '8000-8999',
        position: { x: 55, y: -48 },
      },
      {
        id: 'v1c',
        name: 'T₇-OS · Memory',
        shortLabel: 'T₇-OS',
        function: 'Memory',
        description: 'Stores historical commercial intelligence and experiential learning for later reuse.',
        cognitiveCenter: 'System 4 · Memory',
        dimension: 'Potential Dimension',
        departments: ['Strategic Memory Archive'],
        accountRange: '8000-8999',
        position: { x: 0, y: 62 },
      },
    ],
  },
  {
    id: 'triad2',
    name: 'Triad 2 · Resource Management',
    shortLabel: 'τ₂',
    accountRange: '2400-7999',
    cognitiveCenter: 'Routine / Resource Capacity',
    dimension: 'Potential Dimension',
    description: 'Optimizes procurement, inventory, and performance analytics to convert potential into capacity.',
    departments: ['Resource Optimization Center', 'Procurement Engineering', 'Performance Analytics'],
    financialSummary: ['Resource Optimization Center: $4,200,000', 'Procurement Engineering: $3,100,000', 'Performance Analytics: $2,300,000'],
    position: { x: 200, y: 500 },
    vertices: [
      {
        id: 'v2a',
        name: 'T₅-OS · Operations',
        shortLabel: 'T₅-OS',
        function: 'Operations',
        description: 'Mobilizes material and operating resources in service of production.',
        cognitiveCenter: 'System 4 · Action',
        dimension: 'Commitment Dimension',
        departments: ['Operations Planning'],
        accountRange: '2400-7999',
        position: { x: -55, y: -48 },
      },
      {
        id: 'v2b',
        name: 'T₂-SA · Innovation',
        shortLabel: 'T₂-SA',
        function: 'Innovation',
        description: 'Channels novel procurement and capacity ideas into executable resource plans.',
        cognitiveCenter: 'System 4 · Creation',
        dimension: 'Potential Dimension',
        departments: ['Procurement Engineering'],
        accountRange: '2400-7999',
        position: { x: 55, y: -48 },
      },
      {
        id: 'v2c',
        name: 'T₈-OS · Performance',
        shortLabel: 'T₈-OS',
        function: 'Performance',
        description: 'Evaluates resource throughput and monitors inventory-health feedback.',
        cognitiveCenter: 'System 4 · Balance',
        dimension: 'Performance Dimension',
        departments: ['Performance Analytics'],
        accountRange: '2400-7999',
        position: { x: 0, y: 62 },
      },
    ],
  },
  {
    id: 'triad3',
    name: 'Triad 3 · Innovation',
    shortLabel: 'τ₃',
    accountRange: '5460/3751',
    cognitiveCenter: 'Idea / Creation',
    dimension: 'Potential Dimension',
    description: 'Directs research, adaptation, and translation of new formulations, products, and strategic possibilities.',
    departments: ['Innovation Laboratory', 'Capability Development Institute', 'Technology Integration Office'],
    financialSummary: ['Innovation Laboratory: $3,500,000', 'Capability Development Institute: $2,200,000', 'Technology Integration Office: $1,320,000'],
    position: { x: 350, y: 760 },
    vertices: [
      {
        id: 'v3a',
        name: 'T₂-OS · Directed',
        shortLabel: 'T₂-OS',
        function: 'Directed Creation',
        description: 'Targets innovation toward business priorities and product development paths.',
        cognitiveCenter: 'System 4 · Creation',
        dimension: 'Potential Dimension',
        departments: ['Innovation Laboratory'],
        accountRange: '5460/3751',
        position: { x: -55, y: -48 },
      },
      {
        id: 'v3b',
        name: 'T₆-SA · Adaptive',
        shortLabel: 'T₆-SA',
        function: 'Adaptive Form',
        description: 'Shapes innovation to fit operational and market constraints.',
        cognitiveCenter: 'System 4 · Corporeal',
        dimension: 'Potential Dimension',
        departments: ['Capability Development Institute'],
        accountRange: '5460/3751',
        position: { x: 55, y: -48 },
      },
      {
        id: 'v3c',
        name: 'T₃-OS · Translation',
        shortLabel: 'T₃-OS',
        function: 'Translation',
        description: 'Converts innovative concepts into transferable systems and execution blueprints.',
        cognitiveCenter: 'System 4 · Transfer',
        dimension: 'Commitment Dimension',
        departments: ['Technology Integration Office'],
        accountRange: '5460/3751',
        position: { x: 0, y: 62 },
      },
    ],
  },
  {
    id: 'triad4',
    name: 'Triad 4 · Operations',
    shortLabel: 'τ₄',
    accountRange: '4000-4999',
    cognitiveCenter: 'Routine / Governance',
    dimension: 'Commitment Dimension',
    description: 'Synchronizes autonomous operations, governance, and the observer layer for reliable execution.',
    departments: ['Operations Command Center', 'Quality Assurance Matrix', 'Process Synchronization Hub'],
    financialSummary: ['Operations Command Center: $3,200,000', 'Quality Assurance Matrix: $2,100,000', 'Process Synchronization Hub: $1,500,000'],
    position: { x: 700, y: 780 },
    vertices: [
      {
        id: 'v4a',
        name: 'T₅-SA · Autonomous',
        shortLabel: 'T₅-SA',
        function: 'Autonomous Action',
        description: 'Executes production and coordination routines with minimal latency.',
        cognitiveCenter: 'System 4 · Action',
        dimension: 'Commitment Dimension',
        departments: ['Operations Command Center'],
        accountRange: '4000-4999',
        position: { x: -55, y: -48 },
      },
      {
        id: 'v4b',
        name: 'T₉-OS · Governance',
        shortLabel: 'T₉-OS',
        function: 'Governance',
        description: 'Maintains policy, hierarchy, and control integrity across operating systems.',
        cognitiveCenter: 'System 4 · Hierarchy',
        dimension: 'Commitment Dimension',
        departments: ['Quality Assurance Matrix'],
        accountRange: '4000-4999',
        position: { x: 55, y: -48 },
      },
      {
        id: 'v4c',
        name: 'UT-20 · Observer',
        shortLabel: 'UT-20',
        function: 'Observer',
        description: 'Acts as the transcendental observer integrating cross-triad signals.',
        cognitiveCenter: 'Universal Term 20',
        dimension: 'Performance Dimension',
        departments: ['Process Synchronization Hub'],
        accountRange: '4000-4999',
        position: { x: 0, y: 62 },
      },
    ],
  },
  {
    id: 'triad5',
    name: 'Triad 5 · Customer Interface',
    shortLabel: 'τ₅',
    accountRange: '2000-3400',
    cognitiveCenter: 'Knowledge / Resonance',
    dimension: 'Performance Dimension',
    description: 'Balances customer resonance, homeostasis, and analytics for demand-side stability.',
    departments: ['Customer Fulfillment Operations', 'Distribution Network Management', 'Service Excellence Center'],
    financialSummary: ['Customer Fulfillment Operations: $1,800,000', 'Distribution Network Management: $1,400,000', 'Service Excellence Center: $700,000'],
    position: { x: 1050, y: 760 },
    vertices: [
      {
        id: 'v5a',
        name: 'T₁-SA · Resonance',
        shortLabel: 'T₁-SA',
        function: 'Resonance',
        description: 'Captures emotional and experiential customer response to product delivery.',
        cognitiveCenter: 'System 5 · Limbic Knowledge',
        dimension: 'Performance Dimension',
        departments: ['Customer Fulfillment Operations'],
        accountRange: '2000-3400',
        position: { x: -55, y: -48 },
      },
      {
        id: 'v5b',
        name: 'T₈-SA · Homeostasis',
        shortLabel: 'T₈-SA',
        function: 'Homeostasis',
        description: 'Stabilizes customer service performance and channel consistency.',
        cognitiveCenter: 'System 5 · Balance',
        dimension: 'Performance Dimension',
        departments: ['Distribution Network Management'],
        accountRange: '2000-3400',
        position: { x: 55, y: -48 },
      },
      {
        id: 'v5c',
        name: 'T₄-OS · Analytics',
        shortLabel: 'T₄-OS',
        function: 'Analytics',
        description: 'Organizes demand-side evidence into actionable service and channel insights.',
        cognitiveCenter: 'System 4 · Organization',
        dimension: 'Commitment Dimension',
        departments: ['Service Excellence Center'],
        accountRange: '2000-3400',
        position: { x: 0, y: 62 },
      },
    ],
  },
  {
    id: 'triad6',
    name: 'Triad 6 · Financial Optimization',
    shortLabel: 'τ₆',
    accountRange: '9000-9999',
    cognitiveCenter: 'Knowledge / Financial Balance',
    dimension: 'Performance Dimension',
    description: 'Optimizes liabilities, value engineering, and financial pattern analysis to sustain performance.',
    departments: ['Revenue Optimization Engine', 'Value Engineering Department', 'Financial Pattern Analysis'],
    financialSummary: ['Revenue Optimization Engine: $4,100,000', 'Value Engineering Department: $3,800,000', 'Financial Pattern Analysis: $2,400,000'],
    position: { x: 1200, y: 500 },
    vertices: [
      {
        id: 'v6a',
        name: 'T₆-OS · Infrastructure',
        shortLabel: 'T₆-OS',
        function: 'Infrastructure',
        description: 'Provides the structural substrate for liabilities, payables, and control accounts.',
        cognitiveCenter: 'System 4 · Corporeal',
        dimension: 'Commitment Dimension',
        departments: ['Revenue Optimization Engine'],
        accountRange: '9000-9999',
        position: { x: -55, y: -48 },
      },
      {
        id: 'v6b',
        name: 'T₇-SA · Wisdom',
        shortLabel: 'T₇-SA',
        function: 'Wisdom',
        description: 'Retains strategic lessons from financial control, optimization, and trade-offs.',
        cognitiveCenter: 'System 5 · Memory',
        dimension: 'Potential Dimension',
        departments: ['Value Engineering Department'],
        accountRange: '9000-9999',
        position: { x: 55, y: -48 },
      },
      {
        id: 'v6c',
        name: 'T₃-SA · Morphogenesis',
        shortLabel: 'T₃-SA',
        function: 'Morphogenesis',
        description: 'Reshapes the financial system in response to market and operating dynamics.',
        cognitiveCenter: 'System 5 · Transfer',
        dimension: 'Potential Dimension',
        departments: ['Financial Pattern Analysis'],
        accountRange: '9000-9999',
        position: { x: 0, y: 62 },
      },
    ],
  },
  {
    id: 'triad7',
    name: 'Triad 7 · Organizational Development',
    shortLabel: 'τ₇',
    accountRange: '3350-4500',
    cognitiveCenter: 'Hierarchy / Integration',
    dimension: 'Commitment Dimension',
    description: 'Leads organizational synthesis, change integration, and the interface with the manufacturing core.',
    departments: ['Organizational Development', 'Change Integration Office', 'Systemic Harmony Council'],
    financialSummary: ['Organizational Development: $2,500,000', 'Change Integration Office: $1,800,000', 'Systemic Harmony Council: $900,000'],
    position: { x: 1050, y: 250 },
    vertices: [
      {
        id: 'v7a',
        name: 'T₉-SA · Leadership',
        shortLabel: 'T₉-SA',
        function: 'Leadership',
        description: 'Sets direction and orchestrates change across the cognitive architecture.',
        cognitiveCenter: 'System 5 · Hierarchy',
        dimension: 'Commitment Dimension',
        departments: ['Organizational Development'],
        accountRange: '3350-4500',
        position: { x: -55, y: -48 },
      },
      {
        id: 'v7b',
        name: 'UT-21 · Synthesizer',
        shortLabel: 'UT-21',
        function: 'Synthesizer',
        description: 'Integrates autonomous and analytical knowledge into strategic coherence.',
        cognitiveCenter: 'Universal Term 21',
        dimension: 'Potential Dimension',
        departments: ['Change Integration Office'],
        accountRange: '3350-4500',
        position: { x: 55, y: -48 },
      },
      {
        id: 'v7c',
        name: 'ℳ-Interface · Interface',
        shortLabel: 'ℳ',
        function: 'Interface',
        description: 'Connects organizational development to the manufacturing core’s transformation engine.',
        cognitiveCenter: 'Core Interface',
        dimension: 'Performance Dimension',
        departments: ['Systemic Harmony Council'],
        accountRange: '3350-4500',
        position: { x: 0, y: 62 },
      },
    ],
  },
];

export const PERIPHERAL_NODES: PeripheralNode[] = [
  {
    id: 'manufacturing-core',
    name: 'Manufacturing Core',
    shortLabel: '𝒩^(21)',
    type: 'core',
    description: 'Central pattern consolidation and transformation hub integrating all triads into a living financial organism.',
    dimension: 'All Dimensions',
    departments: ['Pattern Consolidation', 'Transformation Catalyst', 'Omnidirectional Connectivity'],
    financialSummary: ['Pattern Consolidation Effects: $1,500,000', 'Null-Center Catalysis: $920,000', 'JAX CEO Synthesis Gains: $1,180,000'],
    position: { x: 700, y: 500 },
  },
  {
    id: 'jax-ceo',
    name: 'JAX CEO',
    shortLabel: '∇^(∞)⊗∂^ω',
    type: 'executive',
    description: 'Cognitive Executive Orchestration system using neural architecture, auto-differentiation, and resource optimization.',
    dimension: 'Executive Integration',
    departments: ['Neural Architecture', 'Auto-differentiation', 'Optimization Algorithms'],
    financialSummary: ['Neural Architecture Layering', 'Gradient-based Optimization', 'Transfer Learning Across Domains'],
    position: { x: 700, y: 500 },
  },
  {
    id: 'deep-tree-echo',
    name: 'Deep Tree Echo',
    shortLabel: 'DTE',
    type: 'ai',
    description: 'Right-hemisphere complementary AI focused on novelty, intuitive patterning, and the potential dimension.',
    dimension: 'Potential Dimension',
    departments: ['Innovation Sensing', 'Pattern Speculation', 'Opportunity Exploration'],
    financialSummary: ['Innovation Potential Steering', 'Research and formulation exploration'],
    position: { x: 430, y: 110 },
  },
  {
    id: 'marduk',
    name: 'Marduk',
    shortLabel: 'MK',
    type: 'ai',
    description: 'Left-hemisphere complementary AI translating creative possibilities into metric structure and committed execution.',
    dimension: 'Commitment Dimension',
    departments: ['Metric Tensor Analysis', 'Blueprint Translation', 'Production Readiness'],
    financialSummary: ['Administrative and operational translation', 'Financial logic and control preparation'],
    position: { x: 970, y: 110 },
  },
];

export const FINANCIAL_DATA: FinancialData = {
  centers: [
    { name: 'C₁: IDEA (5460/000)', total: 7020000, items: [
      { name: 'Research & Development', value: 3500000 },
      { name: 'Strategic Vision', value: 2200000 },
      { name: 'Deep Tree Echo Systems', value: 1320000 },
    ] },
    { name: 'C₂: KNOWLEDGE (Cerebral) (8000-9999)', total: 10300000, items: [
      { name: 'Financial Intelligence', value: 8420000 },
      { name: 'Analytical Processing', value: 1230000 },
      { name: 'Marduk Logic Systems', value: 650000 },
    ] },
    { name: 'C₃: KNOWLEDGE (Limbic) (3050-3051)', total: 3900000, items: [
      { name: 'Marketing & Brand Value', value: 3050000 },
      { name: 'Customer Experience Assets', value: 510000 },
      { name: 'Emotional Intelligence Reserves', value: 340000 },
    ] },
    { name: 'C₄: ROUTINE (4000-4999)', total: 6800000, items: [
      { name: 'Operational Processes', value: 4850000 },
      { name: 'Standardized Workflows', value: 1200000 },
      { name: 'Synchronization Mechanisms', value: 750000 },
    ] },
    { name: 'C₅: FORM (6000-7999)', total: 15400000, items: [
      { name: 'Physical Assets', value: 6500000 },
      { name: 'Digital Manifestations', value: 1200000 },
      { name: 'Inventory Embodiment', value: 7700000 },
    ] },
  ],
  eliminations: [
    { name: 'τ₁ ↔ τ₅ (Market-Customer Interface)', value: -1200000 },
    { name: 'τ₂ ↔ τ₆ (Resource-Financial)', value: -1800000 },
    { name: 'τ₃ ↔ τ₇ (Innovation-Development)', value: -950000 },
  ],
  adjustments: [
    { name: 'Field Coherence Premium (>0.85)', value: 2100000 },
    { name: 'Metacycle Phase Alignment', value: 850000 },
    { name: 'Universal Term Integration', value: 650000 },
  ],
  transformations: [
    { name: 'Pattern Consolidation Effects', value: 1500000 },
    { name: 'Null-Center Catalysis', value: 920000 },
    { name: 'JAX CEO Synthesis Gains', value: 1180000 },
  ],
};

export const FLOWS: Flow[] = [
  { id: 'flow-supply-1', from: 'triad1', to: 'triad2', type: 'supply', label: '∇¹' },
  { id: 'flow-supply-2', from: 'triad2', to: 'triad7', type: 'supply', label: '∇²' },
  { id: 'flow-supply-3', from: 'triad7', to: 'manufacturing-core', type: 'supply', label: '∇³' },
  { id: 'flow-demand-1', from: 'triad5', to: 'triad6', type: 'demand', label: '∂¹' },
  { id: 'flow-demand-2', from: 'triad6', to: 'triad4', type: 'demand', label: '∂²' },
  { id: 'flow-demand-3', from: 'triad4', to: 'manufacturing-core', type: 'demand', label: '∂³' },
  { id: 'flow-innovation-1', from: 'triad3', to: 'triad4', type: 'innovation', label: 'Ω¹', bidirectional: true },
  { id: 'flow-innovation-2', from: 'triad4', to: 'manufacturing-core', type: 'innovation', label: 'Ω²', bidirectional: true },
  { id: 'flow-innovation-3', from: 'manufacturing-core', to: 'triad3', type: 'innovation', label: 'Ω³', bidirectional: true },
  { id: 'flow-complementary-1', from: 'triad1', to: 'triad5', type: 'complementary', label: 'Ψ^(1,5)', bidirectional: true },
  { id: 'flow-complementary-2', from: 'triad2', to: 'triad6', type: 'complementary', label: 'Ψ^(2,6)', bidirectional: true },
  { id: 'flow-complementary-3', from: 'triad3', to: 'triad7', type: 'complementary', label: 'Ψ^(3,7)', bidirectional: true },
  { id: 'flow-ai-1', from: 'deep-tree-echo', to: 'v4c', type: 'ai', label: 'τ₃ ⊕ τ₁ → UT₂₀' },
  { id: 'flow-ai-2', from: 'marduk', to: 'v7b', type: 'ai', label: 'τ₄ ⊕ τ₆ → UT₂₁' },
];

export const SIMULATION_STEPS: SimulationStep[] = [
  {
    id: 'step-1',
    title: 'Step 1 · Perception',
    description: 'Triad 1 reads market signals and updates the architecture’s perceptual baseline.',
    activeNodeIds: ['triad1'],
    activeFlowIds: ['flow-supply-1', 'flow-complementary-1'],
  },
  {
    id: 'step-2',
    title: 'Step 2 · Organization of Resources',
    description: 'Triad 2 converts demand signals into capacity, inventory, and procurement commitments.',
    activeNodeIds: ['triad2'],
    activeFlowIds: ['flow-supply-1', 'flow-supply-2', 'flow-complementary-2'],
  },
  {
    id: 'step-3',
    title: 'Step 3 · Creative Translation',
    description: 'Triad 3 transforms novel ideas into adaptive product and technology pathways.',
    activeNodeIds: ['triad3', 'deep-tree-echo'],
    activeFlowIds: ['flow-innovation-1', 'flow-ai-1'],
  },
  {
    id: 'step-4',
    title: 'Step 4 · Operational Commitment',
    description: 'Triad 4 synchronizes autonomous execution, governance, and observer feedback.',
    activeNodeIds: ['triad4'],
    activeFlowIds: ['flow-innovation-1', 'flow-demand-3'],
  },
  {
    id: 'step-5',
    title: 'Step 5 · Customer Resonance',
    description: 'Triad 5 measures customer response and keeps downstream service channels in homeostasis.',
    activeNodeIds: ['triad5'],
    activeFlowIds: ['flow-demand-1', 'flow-complementary-1'],
  },
  {
    id: 'step-6',
    title: 'Step 6 · Financial Balance',
    description: 'Triad 6 rebalances obligations, value engineering, and financial pattern intelligence.',
    activeNodeIds: ['triad6', 'marduk'],
    activeFlowIds: ['flow-demand-1', 'flow-demand-2', 'flow-ai-2'],
  },
  {
    id: 'step-7',
    title: 'Step 7 · Synthesis and Reintegration',
    description: 'Triad 7 and the Manufacturing Core consolidate the metacycle into coordinated change.',
    activeNodeIds: ['triad7', 'manufacturing-core', 'jax-ceo'],
    activeFlowIds: ['flow-supply-3', 'flow-innovation-2', 'flow-complementary-3'],
  },
];

const triadNodes = TRIADS.map((triad) => ({
  id: triad.id,
  x: triad.position.x,
  y: triad.position.y,
}));

const vertexNodes = TRIADS.flatMap((triad) => triad.vertices.map((vertex) => ({
  id: vertex.id,
  x: triad.position.x + vertex.position.x,
  y: triad.position.y + vertex.position.y,
})));

const peripheralNodes = PERIPHERAL_NODES.map((node) => ({
  id: node.id,
  x: node.position.x,
  y: node.position.y,
}));

export const NODE_POSITIONS = [...triadNodes, ...vertexNodes, ...peripheralNodes].reduce<Record<string, { x: number; y: number }>>((acc, node) => {
  acc[node.id] = { x: node.x, y: node.y };
  return acc;
}, {});

export const VALID_NODE_IDS = new Set(Object.keys(NODE_POSITIONS));
