import React, { useState, useRef } from 'react';
import { motion, AnimatePresence } from 'framer-motion';
// 만약 프로젝트에서 ES modules나 Webpack/Vite를 쓴다면 logic 가져오기
import { getDashboardCardConfigs, calculateTooltipPosition } from './serviceCombinationLogic';

/**
 * Service Combination Analysis Dashboard Component
 * Data Flow: Logic Module -> React Component (UI)
 */
export default function ServiceCombinationAnalysis() {
    // 순수 함수로부터 카드 메타데이터 및 카운트/사용자 배열 가져오기
    const cardsConfigs = getDashboardCardConfigs();

    // 툴팁 상태 관리 (호버된 카드 ID, 커서 좌표)
    const [hoveredCard, setHoveredCard] = useState(null);
    const [cursorPos, setCursorPos] = useState({ x: 0, y: 0 });

    // 툴팁 닫힘 지연을 위한 Ref (마우스가 카드에서 툴팁으로 이동할 시간 확보)
    const closeTimeoutRef = useRef(null);

    // 카드 마우스 진입 핸들러: 위치 고정 및 툴팁 표시
    const handleCardEnter = (cardId, e) => {
        if (closeTimeoutRef.current) clearTimeout(closeTimeoutRef.current);

        const pos = calculateTooltipPosition(
            e.clientX,
            e.clientY,
            window.innerWidth,
            window.innerHeight,
            300,  // Tooltip estimated width
            400,  // Tooltip max height
            20    // Offset
        );
        setCursorPos(pos);
        setHoveredCard(cardId);
    };

    // 카드/툴팁 마우스 이탈 핸들러
    const handleMouseLeave = () => {
        closeTimeoutRef.current = setTimeout(() => {
            setHoveredCard(null);
        }, 200); // 200ms 여유 시간 제공
    };

    // 현재 호버된 카드의 데이터
    const activeData = hoveredCard ? cardsConfigs.find(c => c.id === hoveredCard) : null;

    return (
        <div className="w-full max-w-6xl mx-auto p-6 font-sans">
            {/* 헤더 섹션 */}
            <div className="mb-6">
                <h1 className="text-2xl font-bold text-gray-900 mb-1">Service Combination Analysis (Upsell Opportunities)</h1>
                <p className="text-gray-500 text-sm">Select a tab or hover over cards to view details</p>
            </div>

            {/* 카드 렌더링 영역 */}
            <div className="flex flex-col md:flex-row gap-5">
                {cardsConfigs.map((card) => (
                    <motion.div
                        key={card.id}
                        onMouseEnter={(e) => handleCardEnter(card.id, e)}
                        onMouseLeave={handleMouseLeave}
                        whileHover={{ scale: 1.02, y: -4 }}
                        transition={{ type: "spring", stiffness: 300, damping: 20 }}
                        className="relative flex-1 bg-white rounded-xl shadow-sm hover:shadow-md cursor-pointer overflow-hidden border border-gray-100 h-44 flex flex-col justify-center"
                    >
                        {/* 왼쪽 고유 색상 보더 효과 */}
                        <div
                            className="absolute left-0 top-0 bottom-0 w-2"
                            style={{ backgroundColor: card.color }}
                        />

                        {/* 카드 내용 */}
                        <div className="pl-8 pr-6 flex items-center justify-between w-full">
                            <h2 className="text-gray-500 text-xs font-bold tracking-wider uppercase">
                                {card.title}
                            </h2>
                            <span className="text-4xl font-extrabold text-gray-900">
                                {card.count}
                            </span>
                        </div>
                    </motion.div>
                ))}
            </div>

            {/* 툴팁 (Portal 없이 직접 absolute/fixed 렌더링) */}
            <AnimatePresence>
                {hoveredCard && activeData && (
                    <motion.div
                        initial={{ opacity: 0, scale: 0.95, y: 10 }}
                        animate={{ opacity: 1, scale: 1, y: 0 }}
                        exit={{ opacity: 0, scale: 0.95, y: 10 }}
                        transition={{ duration: 0.15, ease: "easeOut" }}
                        onMouseEnter={() => { if (closeTimeoutRef.current) clearTimeout(closeTimeoutRef.current); }}
                        onMouseLeave={handleMouseLeave}
                        style={{
                            left: cursorPos.x,
                            top: cursorPos.y,
                        }}
                        className="fixed z-50 w-72 bg-white rounded-lg shadow-2xl border border-gray-100 overflow-hidden pointer-events-auto"
                    >
                        {/* Tooltip Header */}
                        <div
                            className="px-4 py-3 border-b flex justify-between items-center bg-gray-50"
                            style={{ borderTop: `4px solid ${activeData.color}` }}
                        >
                            <span className="font-semibold text-gray-800 text-sm">{activeData.title}</span>
                            <span className="bg-gray-200 text-gray-700 text-xs px-2 py-1 rounded-full font-bold">
                                {activeData.count} Users
                            </span>
                        </div>

                        {/* Tooltip List (Scrollable) */}
                        <div
                            className="max-h-64 overflow-y-auto px-2 py-2 thin-scrollbar overscroll-y-contain"
                            onWheel={(e) => e.stopPropagation()}
                        >
                            {activeData.users.length > 0 ? (
                                <ul className="space-y-1">
                                    {activeData.users.map((user, idx) => (
                                        <motion.li
                                            key={user.id}
                                            initial={{ opacity: 0, x: -10 }}
                                            animate={{ opacity: 1, x: 0 }}
                                            transition={{ delay: idx * 0.02, duration: 0.2 }}
                                            className="px-3 py-2 hover:bg-gray-50 rounded-md text-sm text-gray-700 font-medium flex items-center"
                                        >
                                            <div
                                                className="w-2 h-2 rounded-full mr-3 opacity-70"
                                                style={{ backgroundColor: activeData.color }}
                                            />
                                            {user.name}
                                        </motion.li>
                                    ))}
                                </ul>
                            ) : (
                                <div className="p-4 text-center text-sm text-gray-400">
                                    No tracking entities found.
                                </div>
                            )}
                        </div>
                    </motion.div>
                )}
            </AnimatePresence>

            {/* 스크롤바 커스텀 스타일 지원 (선택사항) */}
            <style jsx="true">{`
                .thin-scrollbar::-webkit-scrollbar {
                    width: 4px;
                }
                .thin-scrollbar::-webkit-scrollbar-track {
                    background: transparent;
                }
                .thin-scrollbar::-webkit-scrollbar-thumb {
                    background-color: #e5e7eb;
                    border-radius: 10px;
                }
            `}</style>
        </div>
    );
}
