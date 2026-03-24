/**
 * serviceCombination Logic 검증용 스크립트
 * 사용자 규칙: 스크립트 작성 로직 검증 최우선 (브라우저 의존성 배제)
 */
const { generateDummyUsers, getDashboardCardConfigs, calculateTooltipPosition } = require('./serviceCombinationLogic');
const assert = require('assert').strict;

// 1. generateDummyUsers 함수 검증
function testGenerateDummyUsers() {
    console.log('[Test 1] generateDummyUsers: 길이 검증');
    assert.strictEqual(generateDummyUsers('activeCustomers', 19).length, 19, '19개 사용자 반환 오류');
    assert.strictEqual(generateDummyUsers('unknown', -1).length, 0, '음수에 대한 방어 실패');
    console.log('✅ 통과');
}

// 2. getDashboardCardConfigs 함수 교차 검증 (데이터 정합성)
function testGetDashboardCardConfigs() {
    console.log('[Test 2] getDashboardCardConfigs: 데이터 정합성 검증');
    const configs = getDashboardCardConfigs();
    
    configs.forEach(config => {
        // count 속성과 생성된 users 배열의 길이가 일치하는지 교차 검증
        assert.strictEqual(config.count, config.users.length, `${config.title} 카운트(${config.count})와 배열 길이(${config.users.length})의 불일치!`);
    });
    console.log('✅ 통과');
}

// 3. calculateTooltipPosition 함수 엣지 케이스 점검
function testCalculateTooltipPosition() {
    console.log('[Test 3] calculateTooltipPosition: 화면 경계 오버플로우 검증');
    
    const winW = 1000;
    const winH = 800;
    const toolW = 200;
    const toolH = 300;
    const offset = 20;

    // 정상 영역
    const normal = calculateTooltipPosition(100, 100, winW, winH, toolW, toolH, offset);
    assert.strictEqual(normal.x, 120, '정상 X 좌표 연산 오류');
    assert.strictEqual(normal.y, 120, '정상 Y 좌표 연산 오류');

    // 우측 화면 오버플로우 시나리오 (마우스가 너비 950일 때, 툴팁이 200이면 화면 넘어감)
    const overflowRight = calculateTooltipPosition(950, 100, winW, winH, toolW, toolH, offset);
    assert.strictEqual(overflowRight.x, 950 - toolW - offset, '우측 오버플로우 방어 로직 오류');
    
    // 하단 화면 오버플로우 시나리오 (마우스 750에 툴팁 높이 300)
    const overflowBottom = calculateTooltipPosition(100, 750, winW, winH, toolW, toolH, offset);
    assert.strictEqual(overflowBottom.y, 750 - toolH - offset, '하단 오버플로우 방어 로직 오류');

    console.log('✅ 통과');
}

try {
    testGenerateDummyUsers();
    testGetDashboardCardConfigs();
    testCalculateTooltipPosition();
    console.log('\n모든 논리 검증이 성공적으로 완료되었습니다.');
} catch(err) {
    console.error('논리 오류 발견:', err.message);
    process.exit(1);
}
