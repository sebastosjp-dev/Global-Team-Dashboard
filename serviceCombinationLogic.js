/**
 * serviceCombinationLogic.js
 * 로직(Pure Function)과 UI(DOM) 엄격 분리 - Node.js 환경에서 독립 검증을 위한 모듈
 */

/**
 * 더미 데이터를 생성하는 순수 함수
 * @param {string} category - 카테고리 식별자 ('activeCustomers', 'multiService', 'upsellTargets')
 * @param {number} count - 생성할 고객 수
 * @returns {Array<{id: string, name: string}>} 고객 목록 배열
 */
function generateDummyUsers(category, count) {
    if (count < 0) return [];
    
    // 카테고리별 접두사 및 회사명 패턴
    const patterns = {
        'activeCustomers': { prefix: 'ac', name: 'Global Corp' },
        'multiService': { prefix: 'ms', name: 'Tech Solutions' },
        'upsellTargets': { prefix: 'up', name: 'Future Enterprise' }
    };
    
    const pattern = patterns[category] || { prefix: 'cm', name: 'Company' };
    
    return Array.from({ length: count }, (_, index) => ({
        id: `${pattern.prefix}-${index + 1}`,
        name: `${pattern.name} ${index + 1}`
    }));
}

/**
 * 대시보드의 카드 데이터 설정을 반환하는 순수 함수
 * 데이터 정합성 유지: 생성된 배열의 길이와 표시할 카운트가 항상 일치하도록 설계
 * @returns {Array<Object>} 카드 설정 목록
 */
function getDashboardCardConfigs() {
    const configs = [
        {
            id: 'active',
            title: 'ACTIVE CUSTOMERS',
            count: 19,
            color: 'rgb(85, 96, 230)', // Blue/Purple (from image)
            category: 'activeCustomers'
        },
        {
            id: 'multi',
            title: 'MULTI-SERVICE',
            count: 14,
            color: 'rgb(41, 182, 115)', // Green (from image)
            category: 'multiService'
        },
        {
            id: 'upsell',
            title: 'UPSELL TARGETS',
            count: 5,
            color: 'rgb(243, 156, 18)', // Orange/Yellow (from image)
            category: 'upsellTargets'
        }
    ];

    // 데이터 교차 검증: 배열 생성 및 실제 데이터 연결
    return configs.map(config => ({
        ...config,
        users: generateDummyUsers(config.category, config.count)
    }));
}

/**
 * 툴팁의 안전한 렌더링 좌표를 계산하는 순수 함수 (마우스 커서 추적용)
 * 화면 밖으로 툴팁이 넘어가는 엣지 케이스 방지
 * @param {number} mouseX - 마우스 X 좌표
 * @param {number} mouseY - 마우스 Y 좌표
 * @param {number} windowWidth - 브라우저 창 너비 (혹은 컨테이너 너비)
 * @param {number} windowHeight - 브라우저 창 높이 (혹은 컨테이너 높이)
 * @param {number} tooltipWidth - 툴팁의 예상 너비
 * @param {number} tooltipHeight - 툴팁의 예상 높이
 * @param {number} offset - 마우스 커서와의 기본 거리
 * @returns {{x: number, y: number}} 계산된 안전한 좌표
 */
function calculateTooltipPosition(mouseX, mouseY, windowWidth, windowHeight, tooltipWidth = 250, tooltipHeight = 300, offset = 20) {
    let x = mouseX + offset;
    let y = mouseY + offset;

    // 우측 화면 밖으로 넘어갈 경우
    if (x + tooltipWidth > windowWidth) {
        x = mouseX - tooltipWidth - offset;
        // 좌측으로도 넘어갈 경우 화면 끝에 고정 (엣지 케이스)
        if (x < 0) x = 0;
    }

    // 하단 화면 밖으로 넘어갈 경우
    if (y + tooltipHeight > windowHeight) {
        y = mouseY - tooltipHeight - offset;
        // 상단으로도 넘어갈 경우 화면 끝에 고정 (엣지 케이스)
        if (y < 0) y = 0;
    }

    return { x, y };
}

module.exports = {
    generateDummyUsers,
    getDashboardCardConfigs,
    calculateTooltipPosition
};
